import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date
from num2words import num2words

# --- 1. دالة تحويل الأرقام ---
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

# إعداد الصفحة
st.set_page_config(page_title="Commune Askaouen - Système PV", layout="wide")

# --- 2. الواجهة الجانبية ---
st.sidebar.header("Membres de la Commission")
p_name = st.sidebar.text_input("Président", "MOHAMED ZILALI")
d_name = st.sidebar.text_input("Directeur du service", "M BAREK BAK")
t_name = st.sidebar.text_input("Technicien", "ABDELLATIF ATTAKY")

st.title("🏛️ نظام استخراج المحاضر - جماعة أسكاون")

# --- 3. التفاصيل الإدارية ---
with st.expander("📝 Détails Administratifs", expanded=True):
    c1, c2 = st.columns(2)
    num_bc = c1.text_input("N° BC", "01/ASK/2026")
    date_pub = c2.date_input("Date de publication", date(2025, 3, 25))
    obj_bc = st.text_area("Objet", "Achat de matériel...")

# --- 4. جدول المتنافسين ---
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

# --- 5. خيارات المحضر ومنطق التحكم ---
c_pv1, c_pv2, c_pv3 = st.columns(3)
pv_num = c_pv1.selectbox("Numéro du PV:", [1, 2, 3, 4, 5, 6])
reunion_date = c_pv3.date_input("Date de la séance", date.today())
reunion_hour = c_pv2.text_input("Heure", "10h00mn")

is_infructueux = False
is_final_attr = False

if pv_num == 6:
    res_6 = st.radio("Résultat du 6éme PV:", ["Attribution", "B.C Infructueux"])
    is_infructueux = (res_6 == "B.C Infructueux")
    is_final_attr = (res_6 == "Attribution")
else:
    # هذا الخيار يجعل أي محضر يعمل كمحضر إسناد نهائي
    is_final_attr = st.checkbox("✅ Est-ce le PV d'attribution finale ? (إسناد نهائي)")

# --- 6. زر إنشاء المحضر ---
if st.button("🚀 إنشاء المحضر"):
    doc = Document()
    # (تنسيق الصفحة والترويسة...)
    section = doc.sections[0]
    section.top_margin, section.bottom_margin = Cm(2), Cm(2)
    
    header = section.header
    htable = header.add_table(1, 2, Inches(6.5))
    htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
    htable.rows[0].cells[1].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
    htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

    doc.add_heading(f"{pv_num}éme Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Objet : {obj_bc}").bold = True
    doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission composée de :")
    doc.add_paragraph(f"- M. {p_name}\n- M. {d_name}\n- M. {t_name}")

    # --- منطق صياغة المحتوى حسب الحالة ---
    
    # الحالة 1: المحضر الأول
    if pv_num == 1:
        doc.add_paragraph("Après vérification، les soumissionnaires sont :")
        tab = doc.add_table(rows=1, cols=3)
        tab.style = 'Table Grid'
        for _, r in data.iterrows():
            row = tab.add_row().cells
            row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
        
        curr_co = data.iloc[0]
        amt_w = format_to_words_fr(curr_co['Montant'])
        doc.add_paragraph(f"\nLa commission invite {curr_co['Nom']} à confirmer son offre ({amt_w}).")
        # --- ضع الجملة الناقصة هنا للمحضر الأول ---
        doc.add_paragraph("الجملة التي كانت تنقصك تضاف هنا...")

    # الحالة 2: إذا تم اختيار "إسناد نهائي" (لأي رقم محضر)
    elif is_final_attr:
        idx = min(pv_num - 1, len(data) - 1)
        curr_co = data.iloc[idx]
        amt_w = format_to_words_fr(curr_co['Montant'])
        doc.add_paragraph(f"La commission constate que la société {curr_co['Nom']} a confirmé son offre.")
        doc.add_paragraph(f"Le président ATTRIBUE le bon de commande à {curr_co['Nom']} pour {curr_co['Montant']} DHS ({amt_w}).").bold = True

    # الحالة 3: غير مثمر
    elif is_infructueux:
        doc.add_paragraph("LA COMMISSION DECLARE QUE CE BON DE COMMANDE EST : INFRUCTUEUX").bold = True

    # الحالة 4: المحاضر الانتقالية (2، 3، 4...)
    else:
        idx = min(pv_num - 1, len(data) - 1)
        prev_co = data.iloc[idx - 1]
        curr_co = data.iloc[idx]
        amt_w = format_to_words_fr(curr_co['Montant'])
        doc.add_paragraph(f"La société {prev_co['Nom']} n'a pas confirmé. Elle est écartée.")
        doc.add_paragraph(f"La commission invite {curr_co['Nom']} ({pv_num}éme) à confirmer son offre ({amt_w}).")

    # التوقيعات
    doc.add_paragraph(f"\nFait à Askaouen, le {reunion_date.strftime('%d/%m/%Y')}").alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # حفظ الملف
    bio = BytesIO()
    doc.save(bio)
    st.download_button(f"📥 تحميل المحضر {pv_num}", bio.getvalue(), f"PV_{pv_num}.docx")
