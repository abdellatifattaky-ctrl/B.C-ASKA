import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. دالة المبالغ ---
def format_to_words_fr(amount_str):
    try:
        val = float(str(amount_str).replace(' ', '').replace(',', ''))
        integer_part = int(val)
        cents = int(round((val - integer_part) * 100))
        words = num2words(integer_part, lang='fr').upper()
        text = f"{words} DHS TTC"
        text += f" ,{num2words(cents, lang='fr').upper()} CENTIMES" if cents > 0 else " ,00CTS"
        return text
    except: return "________________"

st.set_page_config(page_title="Commune Askaouen - Système PV", layout="wide")

# --- 2. إدارة الشعار واللجنة ---
if 'logo_data' not in st.session_state: st.session_state['logo_data'] = None

st.sidebar.header("🖼️ الهوية واللجنة")
uploaded_logo = st.sidebar.file_uploader("ارفع الشعار", type=['png', 'jpg'])
if uploaded_logo: st.session_state['logo_data'] = uploaded_logo.getvalue()

all_members = [
    {"name": "MOHAMED ZILALI", "role": "Président de la commission"},
    {"name": "M BAREK BAK", "role": "Directeur du service"},
    {"name": "ATTAKY ABDELLATIF", "role": "Technicien de la commune"},
    {"name": "NOUREDDIN SALHI", "role": "Service des dépenses"},
    {"name": "FAYSSAL KADRI", "role": "Technicien à la commune"}
]
selected_members = [m for m in all_members if st.sidebar.checkbox(m['name'], value=True)]

st.title("🏛️ نظام استخراج المحاضر - جماعة أسكاون")

# --- 3. البيانات الإدارية ---
with st.expander("📝 Détails Administratifs", expanded=True):
    c1, c2 = st.columns(2)
    num_bc = c1.text_input("N° BC", "01/ASK/2025")
    date_pub = c2.date_input("Date de publication", date(2025, 3, 25))
    obj_bc = st.text_area("Objet", "Location d’une Tractopelle pour les travaux divers.")

st.subheader("📊 Liste des concurrents")
df_init = pd.DataFrame([
    {"Rang": 1, "Nom": "STE OUBRAIM SARL", "Montant": "69840.00"},
    {"Rang": 2, "Nom": "DECO GRC", "Montant": "93120.00"},
    {"Rang": 3, "Nom": "AIT MOUMOU REALISATION ET CONSTRUCTION", "Montant": "102432.00"},
    {"Rang": 4, "Nom": "KADEM SARL", "Montant": "111744.00"},
    {"Rang": 5, "Nom": "TOUZANI 2ZD", "Montant": "114072.00"}
])
data = st.data_editor(df_init, use_container_width=True)

c_v1, c_v2, c_v3 = st.columns(3)
pv_num = c_v1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_v3.date_input("Date de la séance", date.today())
reunion_hour = c_v2.text_input("Heure", "12h00mn")
next_meeting = st.date_input("Date du prochain RDV", reunion_date + timedelta(days=1))
is_final_attr = st.checkbox("✅ إسناد نهائي (Attribution Finale)")

# --- 4. إنشاء المستند ---
if st.button("🚀 إنشاء المحضر النهائي"):
    doc = Document()
    section = doc.sections[0]
    section.left_margin, section.right_margin = Cm(2.5), Cm(2)
    
    # --- الترويسة ---
    header = section.header
    htable = header.add_table(1, 3, Inches(6.5))
    htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
    if st.session_state['logo_data']:
        logo_run = htable.rows[0].cells[1].paragraphs[0].add_run()
        logo_run.add_picture(BytesIO(st.session_state['logo_data']), width=Cm(1.8))
        htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    htable.rows[0].cells[2].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
    htable.rows[0].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # --- العنوان ---
    pv_label = "1ére" if pv_num == 1 else f"{pv_num}éme"
    doc.add_heading(f"{pv_label} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Objet : {obj_bc}").bold = True
    doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission d’ouverture des plis composée comme suit :")
    for m in selected_members:
        doc.add_paragraph(f"- {m['name']} : {m['role']}")

    doc.add_paragraph(f"\nS’est réunie في salle de réunion de la commune بشأن avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')}... ayant pour objet : {obj_bc}")

    idx = min(pv_num - 1, len(data) - 1)
    curr_co = data.iloc[idx]
    amt_words = format_to_words_fr(curr_co['Montant'])

    if pv_num == 1:
        doc.add_paragraph("Les soumissionnaires qui ont déposés leurs offres électroniquement sont :")
        tab = doc.add_table(rows=1, cols=3); tab.style = 'Table Grid'
        hdr = tab.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = 'Classement', 'Concurrent', 'Total TTC'
        for _, r in data.iterrows():
            row = tab.add_row().cells
            row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
        doc.add_paragraph(f"\nFormat papier : Néant.\nLe président invite {curr_co['Nom']} est le moins disant pour un montant de {curr_co['Montant']} Dhs TTC ({amt_words}) à confirmer son offre، suspend la séance et fixe un rendez-vous le {next_meeting.strftime('%d/%m/%Y')} ou sur invitation.")

    elif is_final_attr:
        doc.add_paragraph(f"Après vérification... la commission constate que la société : {curr_co['Nom']} A CONFIRMÉ son offre.")
        p_res = doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le BC à {curr_co['Nom']} لبلغ {curr_co['Montant']} Dhs TTC {amt_words}.")
        p_res.bold = True

    else:
        prev_co = data.iloc[idx - 1]
        doc.add_paragraph(f"تبين أن شركة {prev_co['Nom']} لم تؤكد عرضها.")
        doc.add_paragraph(f"الرئيس يستدعي الشركة التالية: {curr_co['Nom']} (رتبة {pv_num}) بمبلغ {curr_co['Montant']} Dhs ({amt_words}) لتأكيد عرضها، RDV le {next_meeting.strftime('%d/%m/%Y')}.")

    # --- جدول التوقيعات (بفراغات واسعة جداً) ---
    doc.add_paragraph(f"\nAskaouen le {reunion_date.strftime('%d/%m/%Y')}")
    sig_tab = doc.add_table(rows=0, cols=2)
    sig_tab.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # إضافة الأسماء في أزواج مع مسافة عمودية كبيرة
    for i in range(0, len(selected_members), 2):
        row_names = sig_tab.add_row()
        row_names.height = Cm(0.8) # ارتفاع لاسم العضو
        cell_left = row_names.cells[0]
        cell_right = row_names.cells[1]
        
        cell_left.text = selected_members[i]['name']
        cell_left.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        if i + 1 < len(selected_members):
            cell_right.text = selected_members[i+1]['name']
            cell_right.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            
        # إضافة سطر فارغ كبير للتوقيع تحت كل اسم
        row_space = sig_tab.add_row()
        row_space.height = Cm(3.0) # <--- هذا الفراغ (3 سم) مخصص للتوقيع والختم

    bio = BytesIO()
    doc.save(bio)
    st.success("✅ تم تحديث التنسيق وتوسيع أماكن التوقيع!")
    st.download_button("📥 تحميل الملف النهائي", bio.getvalue(), f"PV_Askaouen_{pv_num}.docx")
