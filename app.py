import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. دالة تحويل الأرقام إلى حروف ---
def format_to_words_fr(amount_str):
    try:
        val = float(str(amount_str).replace(' ', '').replace(',', ''))
        integer_part = int(val)
        cents = int(round((val - integer_part) * 100))
        words = num2words(integer_part, lang='fr').upper()
        text = words + " DHS TTC"
        if cents > 0:
            cents_words = num2words(cents, lang='fr').upper()
            text += f" ,{cents_words} CENTIMES"
        else:
            text += " ,00CTS"
        return text
    except:
        return "________________"

# --- 2. إعدادات الصفحة ---
st.set_page_config(page_title="Commune Askaouen - Système PV", layout="wide")

# --- 3. إدارة أعضاء اللجنة (Checkboxes) ---
st.sidebar.header("👥 Membres de la Commission")
st.sidebar.write("Sélectionnez les membres présents :")

# تعريف البيانات الأساسية للأعضاء
all_members = [
    {"id": "p", "name": "MOHAMED ZILALI", "role": "Président de la commission"},
    {"id": "d", "name": "M BAREK BAK", "role": "Directeur du service"},
    {"id": "t1", "name": "ATTAKY ABDELLATIF", "role": "Technicien de la commune"},
    {"id": "s", "name": "NOUREDDIN SALHI", "role": "Service des dépenses"},
    {"id": "t2", "name": "FAYSSAL KADRI", "role": "Technicien à la commune"}
]

selected_members = []
for member in all_members:
    if st.sidebar.checkbox(f"{member['name']}", value=True):
        # السماح بتعديل الاسم أو الدور حتى بعد الاختيار
        m_name = st.sidebar.text_input(f"Nom ({member['id']})", member['name'], label_visibility="collapsed")
        selected_members.append({"name": m_name, "role": member['role']})

st.title("🏛️ نظام استخراج المحاضر - جماعة أسكاون")

# --- 4. التفاصيل الإدارية ---
with st.expander("📝 Détails Administratifs", expanded=True):
    c1, c2 = st.columns(2)
    num_bc = c1.text_input("N° BC", "01/ASK/2025")
    date_pub = c2.date_input("Date de publication", date(2025, 3, 25))
    obj_bc = st.text_area("Objet", "Location d’une Tractopelle pour les travaux divers.")

# --- 5. جدول المتنافسين ---
st.subheader("📊 Liste des concurrents")
df_init = pd.DataFrame([
    {"Rang": 1, "Nom": "STE OUBRAIM SARL", "Montant": "69840.00"},
    {"Rang": 2, "Nom": "DECO GRC", "Montant": "93120.00"},
    {"Rang": 3, "Nom": "AIT MOUMOU REALISATION ET CONSTRUCTION", "Montant": "102432.00"},
    {"Rang": 4, "Nom": "KADEM SARL", "Montant": "111744.00"},
    {"Rang": 5, "Nom": "TOUZANI 2ZD", "Montant": "114072.00"}
])
data = st.data_editor(df_init, use_container_width=True)

st.divider()

# --- 6. خيارات المحضر والمواعيد ---
c_pv1, c_pv2, c_pv3 = st.columns(3)
pv_num = c_pv1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_pv3.date_input("Date de la séance", date.today())
reunion_hour = c_pv2.text_input("Heure", "12h00mn")

next_meeting_default = reunion_date + timedelta(days=1)
next_meeting = st.date_input("Date du prochain rendez-vous (J+1)", next_meeting_default)

is_final_attr = False
is_infructueux = False

if pv_num == 6:
    res_6 = st.radio("Résultat du 6éme PV:", ["Attribution", "B.C Infructueux"])
    is_infructueux = (res_6 == "B.C Infructueux")
    is_final_attr = (res_6 == "Attribution")
else:
    is_final_attr = st.checkbox("✅ Est-ce le PV d'attribution finale ?")

# --- 7. إنشاء مستند Word ---
if st.button("🚀 إنشاء المحضر"):
    if not selected_members:
        st.error("المرجو اختيار عضو واحد على الأقل في اللجنة!")
    else:
        doc = Document()
        section = doc.sections[0]
        section.top_margin, section.bottom_margin = Cm(2), Cm(2)

        header = section.header
        htable = header.add_table(1, 2, Inches(6.5))
        htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
        htable.rows[0].cells[1].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
        htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

        pv_label = "1ére" if pv_num == 1 else f"{pv_num}éme"
        doc.add_heading(f"{pv_label} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph(f"Objet : {obj_bc}").bold = True
        doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission composée de :")
        
        # كتابة الأعضاء المختارين فقط
        for m in selected_members:
            doc.add_paragraph(f"- {m['name']} : {m['role']}")

        doc.add_paragraph(f"\nS’est réunie في salle de réunion de la commune concernant l’avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')} sur le portail des marchés publics، ayant pour objet : {obj_bc}")

        idx = min(pv_num - 1, len(data) - 1)
        curr_co = data.iloc[idx]
        amt_words = format_to_words_fr(curr_co['Montant'])

        if pv_num == 1:
            doc.add_paragraph("Les soumissionnaires qui ont déposés leurs offres électroniquement sont :")
            tab = doc.add_table(rows=1, cols=3)
            tab.style = 'Table Grid'
            hdr = tab.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Classement', 'Nom des concurrents', 'Total TTC'
            for _, r in data.iterrows():
                row = tab.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
            
            doc.add_paragraph("\nFormat papier : Néant.")
            doc.add_paragraph(f"Le président invite la société : {curr_co['Nom']} (moins disant) pour un montant de {curr_co['Montant']} Dhs TTC ({amt_words}) à confirmer son offre، suspend la séance et fixe un rendez-vous le {next_meeting.strftime('%d/%m/%Y')} ou sur invitation.")

        elif is_infructueux:
            doc.add_paragraph(f"La commission constate que la société {curr_co['Nom']} n’a pas confirmé son offre.")
            p_inf = doc.add_paragraph("\nPAR CONSEQUENT, LA COMMISSION DECLARE QUE CE BON DE COMMANDE EST :")
            p_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("INFRUCTUEUX").alignment = WD_ALIGN_PARAGRAPH.CENTER

        elif is_final_attr:
            doc.add_paragraph(f"Après vérification... la commission constate que la société : {curr_co['Nom']} A CONFIRMÉ son offre.")
            doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le BC à {curr_co['Nom']} pour {curr_co['Montant']} Dhs TTC {amt_words}.").bold = True

        else:
            prev_co = data.iloc[idx - 1]
            doc.add_paragraph(f"Après vérification... {prev_co['Nom']} n’a pas confirmé.")
            doc.add_paragraph(f"Le président invite {curr_co['Nom']} ({pv_num}éme) بمبلغ {curr_co['Montant']} Dhs TTC {amt_words} à confirmer، RDV le {next_meeting.strftime('%d/%m/%Y')}.")

        # التوقيعات (تلقائية بناءً على المختارين)
        doc.add_paragraph(f"\nAskaouen le {reunion_date.strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
        sig_tab = doc.add_table(rows=(len(selected_members) // 2) + 1, cols=2)
        for i, m in enumerate(selected_members):
            row_idx = i // 2
            col_idx = i % 2
            sig_tab.cell(row_idx, col_idx).text = m['name']
            sig_tab.cell(row_idx, col_idx).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

        bio = BytesIO()
        doc.save(bio)
        st.success("✅ تم توليد المحضر بالأعضاء المختارين!")
        st.download_button(f"📥 تحميل PV {pv_num}", bio.getvalue(), f"PV_{pv_num}.docx")
