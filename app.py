import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date
from num2words import num2words

# دالة تحويل الأرقام إلى حروف بالفرنسية
def format_to_words_fr(amount_str):
    try:
        val = float(str(amount_str).replace(' ', '').replace(',', ''))
        words = num2words(val, lang='fr').upper()
        cents = int(round((val - int(val)) * 100))
        text = f"{words} DIRHAMS"
        if cents > 0:
            text += f" ET {num2words(cents, lang='fr').upper()} CENTIMES"
        else:
            text += " ,00CTS"
        return text
    except:
        return "________________"

st.set_page_config(page_title="Commune Askaouen - Système PV", layout="wide")

# أعضاء اللجنة في الشبيكة الجانبية
st.sidebar.header("Membres de la Commission")
p_name = st.sidebar.text_input("Président", "MOHAMED ZILALI")
d_name = st.sidebar.text_input("Directeur du service", "M BAREK BAK")
t_name = st.sidebar.text_input("Technicien", "ABDELLATIF ATTAKY")

st.title("🏛️ نظام استخراج المحاضر - جماعة أسكاون")

# تفاصيل إدارية
with st.expander("📝 Détails Administratifs", expanded=True):
    c1, c2 = st.columns(2)
    num_bc = c1.text_input("N° BC", "01/ASK/2026")
    date_pub = c2.date_input("Date de publication", date(2025, 3, 25))
    obj_bc = st.text_area("Objet", "Achat de matériel...")

# جدول المتنافسين
st.subheader("📊 Liste des concurrents")
df_init = pd.DataFrame([
    {"Rang": 1, "Nom": "STE OUBRAIM SARL", "Montant": "69840.00"},
    {"Rang": 2, "Nom": "DECO GRC", "Montant": "93120.00"},
    {"Rang": 3, "Nom": "AIT MOUMOU REALISATION", "Montant": "102432.00"},
    {"Rang": 4, "Nom": "KADEM SARL", "Montant": "111744.00"},
    {"Rang": 5, "Nom": "TOUZANI 2ZD", "Montant": "114072.00"}
])
data = st.data_editor(df_init, use_container_width=True)

st.divider()

# خيارات المحضر
c_pv1, c_pv2, c_pv3 = st.columns(3)
pv_num = c_pv1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_pv3.date_input("Date de la séance", date.today())
reunion_hour = st.text_input("Heure", "10h00mn")

is_infructueux = False
is_final_attr = False

# منطق تحديد حالة المحضر
if pv_num == 6:
    res_6 = st.radio("Résultat du 6éme PV:", ["Attribution (إسناد الشركة 5)", "B.C Infructueux (غير مثمر)"])
    is_infructueux = (res_6 == "B.C Infructueux (غير مثمر)")
    is_final_attr = (res_6 == "Attribution (إسناد الشركة 5)")
else:
    is_final_attr = c_pv2.checkbox("✅ Est-ce le PV d'attribution finale ?")

# زر إنشاء المحضر
if st.button("🚀 إنشاء المحضر"):
    doc = Document()
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(2), Cm(2)
    section.left_margin, section.right_margin = Cm(2.5), Cm(2)

    # الترويسة (Header)
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
    htable.rows[0].cells[1].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
    htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_paragraph("\n")
    doc.add_heading(f"{pv_num}éme Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(f"Objet : {obj_bc}").bold = True
    doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission d’ouverture des plis composée comme suit :")
    doc.add_paragraph(f"- M. {p_name} : Président de la commission\n- M. {d_name} : Directeur du service\n- M. {t_name} : Technicien de la commune")

    doc.add_paragraph(f"S’est réunie dans la salle de réunion de la commune sur invitation du président concernant l’avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')} sur le portail des marchés publics, en application des dispositions de l’article 91 du décret n° 2-22-431 (8 mars 2023) relatif aux marchés publics.")

    if pv_num == 1:
        doc.add_paragraph("Après vérification du portail des marchés publics, les soumissionnaires qui ont déposé leurs offres de prix électroniquement sont :")
        tab = doc.add_table(rows=1, cols=3)
        tab.style = 'Table Grid'
        hdr = tab.rows[0].cells
        hdr[0].text, hdr[1].text, hdr[2].text = 'Rang', 'Concurrent', 'Montant TTC'
        for _, r in data.iterrows():
            row = tab.add_row().cells
            row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"

        curr_company = data.iloc[0]['Nom']
        curr_amount = data.iloc[0]['Montant']
        amt_w = format_to_words_fr(curr_amount)
        doc.add_paragraph(f"\nAprès examen des offres, le président de la commission invite la société : {curr_company} qui est le moins disant pour un montant de {curr_amount} Dhs TTC ({amt_w}) à confirmer son offre par lettre de confirmation.")

    else:
        # المحاضر من 2 إلى 6
        idx = pv_num - 1 if pv_num <= 5 else 4
        prev_idx = idx - 1
        prev_company = data.iloc[prev_idx]['Nom']
        curr_company = data.iloc[idx]['Nom']
        curr_amount = data.iloc[idx]['Montant']
        amt_w = format_to_words_fr(curr_amount)

        if is_infructueux:
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {curr_company} n’a pas confirmé son offre par lettre de confirmation.")
            p_inf = doc.add_paragraph("\nPAR CONSEQUENT, LA COMMISSION DECLARE QUE CE BON DE COMMANDE EST :")
            p_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            res_inf = doc.add_paragraph("INFRUCTUEUX")
            res_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            res_inf.bold = True
            res_inf.runs[0].font.size = Pt(16)

        elif is_final_attr:
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {curr_company} a confirmé son offre par lettre de confirmation.")
            p_res = doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le bon de commande à la société {curr_company} pour un montant de : {curr_amount} Dhs TTC ({amt_w}).")
            p_res.bold = True

        else:
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {prev_company} n’a pas confirmé son offre par lettre de confirmation.")
            doc.add_paragraph(f"Après écartement de la société {prev_company}, le président de la commission invite la société : {curr_company} qui est classé le {pv_num}éme pour un montant de {curr_amount} Dhs TTC ({amt_w}) à confirmer son offre par lettre de confirmation.")

    # التوقيعات في أسفل الصفحة
    doc.add_paragraph(f"\nFait à Askaouen, le {reunion_date.strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sig_tab = doc.add_table(rows=2, cols=3)
    sig_tab.rows[0].cells[0].text = "Le Président"
    sig_tab.rows[0].cells[1].text = "Le Directeur"
    sig_tab.rows[0].cells[2].text = "Le Technicien"
    sig_tab.rows[1].cells[0].text = p_name
    sig_tab.rows[1].cells[1].text = d_name
    sig_tab.rows[1].cells[2].text = t_name
    
    for r in sig_tab.rows:
        for c in r.cells:
            c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    # حفظ وتحميل الملف
    bio = BytesIO()
    doc.save(bio)
    st.download_button(f"📥 تحميل المحضر رقم {pv_num}", bio.getvalue(), f"PV_{pv_num}_Askaouen.docx")
