import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# دالة تحويل الأرقام إلى حروف بالفرنسية (مع إضافة CTS)
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

st.set_page_config(page_title="Commune Askaouen - Système PV", layout="wide")

# --- الشبيكة الجانبية ---
st.sidebar.header("Membres de la Commission")
p_name = st.sidebar.text_input("Président", "MOHAMED ZILALI")
d_name = st.sidebar.text_input("Directeur du service", "M BAREK BAK")
t_name = st.sidebar.text_input("Technicien", "ABDELLATIF ATTAKY")

st.title("🏛️ نظام استخراج المحاضر - جماعة أسكاون")

# --- التفاصيل الإدارية ---
with st.expander("📝 Détails Administratifs", expanded=True):
    c1, c2 = st.columns(2)
    num_bc = c1.text_input("N° BC", "01/ASK/2025")
    date_pub = c2.date_input("Date de publication", date(2025, 3, 25))
    obj_bc = st.text_area("Objet", "Location d’une Tractopelle pour les travaux divers.")

# --- جدول المتنافسين ---
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

# --- خيارات المحضر ---
c_pv1, c_pv2, c_pv3 = st.columns(3)
pv_num = c_pv1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_pv3.date_input("Date de la séance", date.today())
reunion_hour = c_pv2.text_input("Heure", "12h00mn")

# خيار الإسناد النهائي (يجعل أي محضر يعمل كمحضر ختامي)
is_final_attr = st.checkbox("✅ Est-ce le PV d'attribution finale ? (إسناد الشركة الحالية)")
next_meeting = st.date_input("Date du prochain rendez-vous (إن وجد)", reunion_date + timedelta(days=1))

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

    # العنوان
    doc.add_heading(f"{pv_num}éme Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Objet : {obj_bc}").bold = True
    
    # أعضاء اللجنة بتنسيق مشابه للجدول
    doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission d’ouverture des plis composée comme suit :")
    sig_list = doc.add_table(rows=3, cols=2)
    sig_list.cell(0,0).text = f"- {p_name}"
    sig_list.cell(0,1).text = "Président de la commission"
    sig_list.cell(1,0).text = f"- {d_name}"
    sig_list.cell(1,1).text = "Directeur du service"
    sig_list.cell(2,0).text = f"- {t_name}"
    sig_list.cell(2,1).text = "Technicien de la commune"

    doc.add_paragraph(f"\nS’est réunie dans la salle de la réunion de la commune sur invitation du président... concernant l’avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')}... ayant pour objet : {obj_bc}")

    # الفهرس الحالي بناءً على رقم المحضر
    idx = min(pv_num - 1, len(data) - 1)
    curr_co = data.iloc[idx]
    amt_w = format_to_words_fr(curr_co['Montant'])

    if pv_num == 1:
        doc.add_paragraph("Les soumissionnaires qui ont déposés leurs offres de prix électroniquement sont :")
        tab = doc.add_table(rows=1, cols=3)
        tab.style = 'Table Grid'
        hdr = tab.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = 'Classement', 'Nom des concurrents', 'Total TTC'
        for _, r in data.iterrows():
            row = tab.add_row().cells
            row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
        
        doc.add_paragraph("\nFormat papier : Néant.")
        doc.add_paragraph(f"Le président... invite la société : {curr_co['Nom']} est le moins disant pour un montant de {curr_co['Montant']} Dhs TTC ({amt_w}) à confirmer son offre, et suspend la séance et fixe un rendez-vous le {next_meeting.strftime('%d/%m/%Y')} ou sur invitation.")

    elif is_final_attr:
        # صياغة المحضر النهائي (مثل المحضر الرابع في مثالك)
        doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société : {curr_co['Nom']} a confirmé son offre par lettre de confirmation.")
        p_fin = doc.add_paragraph(f"Le président de la commission valide la confirmation et attribue le bon de commande à la société {curr_co['Nom']} pour un montant de : {curr_co['Montant']} Dhs TTC {amt_w}.")
        p_fin.bold = True

    else:
        # صياغة المحاضر الانتقالية (2، 3، إلخ)
        prev_co = data.iloc[idx - 1]
        doc.add_paragraph(f"Après vérification... la commission constate que la société {prev_co['Nom']} n’a pas confirmé son offre par lettre de confirmation.")
        doc.add_paragraph(f"Après écartement de la société {prev_co['Nom']} le président de la commission invite la société : {curr_co['Nom']} qui est classé le {pv_num}éme pour un montant de {curr_co['Montant']} Dhs TTC {amt_w} à confirmer son offre, et suspend la séance et fixe un rendez-vous le {next_meeting.strftime('%d/%m/%Y')} ou sur invitation.")

    # التوقيعات الختامية
    doc.add_paragraph(f"\nAskaouen le {reunion_date.strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_tab = doc.add_table(rows=2, cols=3)
    sig_tab.cell(0,0).text = "MOHAMED ZILALI"
    sig_tab.cell(0,1).text = "M BAREK BAK"
    sig_tab.cell(0,2).text = "ATTAKY ABDELLATIF"
    for cell in sig_tab.rows[0].cells:
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    bio = BytesIO()
    doc.save(bio)
    st.success("✅ تم توليد المحضر بنجاح!")
    st.download_button(f"📥 تحميل المحضر رقم {pv_num}", bio.getvalue(), f"PV_{pv_num}.docx")
