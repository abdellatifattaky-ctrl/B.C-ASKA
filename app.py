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

# --- 3. الواجهة الجانبية (إضافة الأعضاء الجدد) ---
st.sidebar.header("Membres de la Commission")
p_name = st.sidebar.text_input("Président", "MOHAMED ZILALI")
d_name = st.sidebar.text_input("Directeur du service", "M BAREK BAK")
t_name = st.sidebar.text_input("Technicien 1", "ATTAKY ABDELLATIF")
s_name = st.sidebar.text_input("Service des dépenses", "NOUREDDIN SALHI")
f_name = st.sidebar.text_input("Technicien 2", "FAYSSAL KADRI")

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

# --- 6. منطق التحكم والمواعيد ---
c_pv1, c_pv2, c_pv3 = st.columns(3)
pv_num = c_pv1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_pv3.date_input("Date de la séance", date.today())
reunion_hour = c_pv2.text_input("Heure", "12h00mn")

# التوقع التلقائي للموعد القادم (يوم + 1)
next_meeting_default = reunion_date + timedelta(days=1)
next_meeting = st.date_input("Date du prochain rendez-vous (J+1)", next_meeting_default)

is_final_attr = False
is_infructueux = False

if pv_num == 6:
    res_6 = st.radio("Résultat du 6éme PV:", ["Attribution (إسناد)", "B.C Infructueux (بدون جدوى)"])
    is_infructueux = (res_6 == "B.C Infructueux (بدون جدوى)")
    is_final_attr = (res_6 == "Attribution (إسناد)")
else:
    is_final_attr = st.checkbox("✅ Est-ce le PV d'attribution finale ?")

# --- 7. إنشاء مستند Word ---
if st.button("🚀 إنشاء المحضر"):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(2), Cm(2)

    # الترويسة
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
    htable.rows[0].cells[1].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
    htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # تصحيح تسمية المحضر الأول (1ére)
    pv_title = f"{pv_num}ére" if pv_num == 1 else f"{pv_num}éme"
    doc.add_heading(f"{pv_title} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Objet : {obj_bc}").bold = True
    doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission composée de :")
    
    # إضافة الأعضاء الخمسة للجنة
    doc.add_paragraph(f"- {p_name} : Président de la commission")
    doc.add_paragraph(f"- {d_name} : Directeur du service")
    doc.add_paragraph(f"- {t_name} : Technicien de la commune")
    doc.add_paragraph(f"- {s_name} : Service des dépenses")
    doc.add_paragraph(f"- {f_name} : Technicien à la commune")

    doc.add_paragraph(f"\nS’est réunie dans la salle de réunion de la commune concernant l’avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')} sur le portail des marchés publics، ayant pour objet : {obj_bc}")

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
        doc.add_paragraph(f"Après vérification... la commission constate que la société {curr_co['Nom']} n’a pas confirmé son offre.")
        p_inf = doc.add_paragraph("\nPAR CONSEQUENT, LA COMMISSION DECLARE QUE CE BON DE COMMANDE EST :")
        p_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        res_t = doc.add_paragraph("INFRUCTUEUX")
        res_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        res_t.bold = True

    elif is_final_attr:
        doc.add_paragraph(f"Après vérification du portail des marchés publics، la commission constate que la société : {curr_co['Nom']} A CONFIRMÉ son offre par lettre de confirmation.")
        p_res = doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le bon de commande à la société {curr_co['Nom']} pour un montant de : {curr_co['Montant']} Dhs TTC {amt_words}.")
        p_res.bold = True

    else:
        prev_co = data.iloc[idx - 1]
        doc.add_paragraph(f"Après vérification... la commission constate que la société {prev_co['Nom']} n’a pas confirmé son offre par lettre de confirmation.")
        doc.add_paragraph(f"Après écartement de la société {prev_co['Nom']} le président invite la société : {curr_co['Nom']} (classé {pv_num}éme) pour un montant de {curr_co['Montant']} Dhs TTC {amt_words} à confirmer son offre، suspend la séance et fixe un rendez-vous le {next_meeting.strftime('%d/%m/%Y')} ou sur invitation.")

    # التوقيعات (تشمل الأعضاء الجدد)
    doc.add_paragraph(f"\nAskaouen le {reunion_date.strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_tab = doc.add_table(rows=2, cols=3)
    sig_tab.cell(0,0).text = p_name
    sig_tab.cell(0,1).text = d_name
    sig_tab.cell(0,2).text = t_name
    sig_tab.cell(1,0).text = s_name
    sig_tab.cell(1,1).text = f_name
    for row in sig_tab.rows:
        for cell in row.cells:
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    bio = BytesIO()
    doc.save(bio)
    st.success(f"✅ تم تحديث اللجنة وتصحيح عنوان المحضر {pv_num}!")
    st.download_button(f"📥 تحميل المحضر رقم {pv_num}", bio.getvalue(), f"PV_{pv_num}.docx")
