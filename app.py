import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. الدوال المساعدة ---
def format_to_words_fr(amount_str):
    try:
        val = float(str(amount_str).replace(' ', '').replace(',', ''))
        integer_part = int(val)
        cents = int(round((val - integer_part) * 100))
        words = num2words(integer_part, lang='fr').upper()
        text = f"{words} DHS TTC"
        text += f" ET {num2words(cents, lang='fr').upper()} CENTIMES" if cents > 0 else " ,00CTS"
        return text
    except: return "________________"

def apply_official_header(doc, logo_data):
    section = doc.sections[0]
    header = section.header
    htable = header.add_table(1, 3, Inches(6.5))
    # يسار - فرنسي
    c_left = htable.rows[0].cells[0].paragraphs[0]
    c_left.text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nPROVINCE DE TAROUDANT\nCOMMUNE D'ASKAOUN"
    c_left.style.font.size = Pt(9)
    # وسط - شعار
    if logo_data:
        try:
            logo_run = htable.rows[0].cells[1].paragraphs[0].add_run()
            logo_run.add_picture(BytesIO(logo_data), width=Cm(1.8))
            htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        except: pass
    # يمين - عربي
    c_right = htable.rows[0].cells[2].paragraphs[0]
    c_right.text = "المملكة المغربية\nوزارة الداخلية\nإقليم تارودانت\nجماعة أسكاون"
    c_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c_right.style.font.size = Pt(9)

# --- 2. إعدادات الجلسة والواجهة ---
st.set_page_config(page_title="Système Intégré Askaouen", layout="wide")

if 'logo_data' not in st.session_state:
    st.session_state['logo_data'] = None

# --- 3. الشريط الجانبي (Logo & Members) ---
st.sidebar.header("⚙️ الإعدادات العامة")
uploaded_logo = st.sidebar.file_uploader("ارفع شعار الجماعة", type=['png', 'jpg', 'jpeg'])
if uploaded_logo:
    st.session_state['logo_data'] = uploaded_logo.getvalue()

st.sidebar.divider()
st.sidebar.header("👥 أعضاء اللجنة")
members_list = [
    {"name": "MOHAMED ZILALI", "role": "Président de la commission"},
    {"name": "M BAREK BAK", "role": "Directeur du service"},
    {"name": "ATTAKY ABDELLATIF", "role": "Technicien de la commune"},
    {"name": "NOUREDDIN SALHI", "role": "Service des dépenses"},
    {"name": "FAYSSAL KADRI", "role": "Technicien à la commune"}
]
selected_members = []
for m in members_list:
    if st.sidebar.checkbox(m['name'], value=True):
        selected_members.append(m)

# --- 4. التبويبات الرئيسية ---
tab_pv, tab_docs, tab_photos = st.tabs(["🏛️ المحاضر (PV 1-6)", "✉️ الوثائق الرسمية (OS)", "📸 صور الأوراش"])

# --- التبويب الأول: المحاضر الرسمية ---
with tab_pv:
    st.subheader("توليد محاضر فتح الأظرفة (1ére - 6éme)")
    with st.expander("📝 معلومات طلب السند (BC)", expanded=True):
        c1, c2 = st.columns(2)
        bc_num = c1.text_input("N° BC", "01/ASK/2026")
        bc_date_pub = c2.date_input("تاريخ النشر", date(2026, 3, 1))
        bc_objet = st.text_area("الموضوع (Objet)", "Location de matériel...")

    st.subheader("📊 قائمة المتنافسين")
    df_init = pd.DataFrame([
        {"Rang": 1, "Nom": "STE OUBRAIM SARL", "Montant": "0.00"},
        {"Rang": 2, "Nom": "DECO GRC", "Montant": "0.00"},
        {"Rang": 3, "Nom": "AIT MOUMOU REALISATION", "Montant": "0.00"}
    ])
    competitors = st.data_editor(df_init, use_container_width=True)

    c3, c4, c5 = st.columns(3)
    pv_index = c3.selectbox("رقم المحضر:", [1, 2, 3, 4, 5, 6])
    session_date = c4.date_input("تاريخ الجلسة", date.today())
    next_date = c5.date_input("الموعد القادم", session_date + timedelta(days=1))
    
    is_final_attr = st.checkbox("✅ إسناد نهائي (Attribution)")

    if st.button("🚀 توليد المحضر الرسمي"):
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        
        pv_label = "1ére" if pv_index == 1 else f"{pv_index}éme"
        doc.add_heading(f"{pv_label} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"Objet : {bc_objet}").bold = True
        doc.add_paragraph(f"Le {session_date.strftime('%d/%m/%Y')}, la commission composée de :")
        for m in selected_members:
            doc.add_paragraph(f"- {m['name']} : {m['role']}")

        doc.add_paragraph(f"\nS’est réunie dans la salle de la réunion de la commune concernant l’avis d’achat du bon de commande n° {bc_num} publié le : {bc_date_pub.strftime('%d/%m/%Y')} على بوابة الصفقات العمومية.")

        idx = min(pv_index - 1, len(competitors) - 1)
        curr = competitors.iloc[idx]
        amt_txt = format_to_words_fr(curr['Montant'])

        if pv_index == 1:
            doc.add_paragraph(f"Le président invite la société : {curr['Nom']} (moins disant) بمبلغ {curr['Montant']} Dhs TTC ({amt_txt}) à confirmer son offre، RDV le {next_date.strftime('%d/%m/%Y')}.")
        elif is_final_attr:
            doc.add_paragraph(f"La commission VALIDE la confirmation et ATTRIBUE le BC à {curr['Nom']} بمبلغ {curr['Montant']} Dhs TTC.")
        else:
            prev = competitors.iloc[idx-1]
            doc.add_paragraph(f"Après vérification... {prev['Nom']} n’a pas confirmé. Le président invite {curr['Nom']} ({pv_index}éme) بمبلغ {curr['Montant']} Dhs à confirmer، RDV le {next_date.strftime('%d/%m/%Y')}.")

        # توقيعات
        sig_tab = doc.add_table(rows=2, cols=2); sig_tab.rows[1].height = Cm(2.5)
        
        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 تحميل المحضر", bio.getvalue(), f"PV_{pv_index}.docx")

# --- التبويب الثاني: OS & Notification ---
with tab_docs:
    st.subheader("توليد الإشعار (Notification) وأمر الخدمة (OS)")
    with st.form("docs_form"):
        f_soc = st.text_input("اسم الشركة الفائزة")
        f_amt = st.text_input("المبلغ الإجمالي (DHS TTC)")
        f_delai = st.number_input("مدة التنفيذ (أيام)", value=10)
        submit_docs = st.form_submit_button("🚀 توليد المستندات")

    if submit_docs:
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        doc.add_heading("LETTRE DE NOTIFICATION", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"\nA la société : {f_soc}\nObjet : Notification du BC n° {bc_num}")
        doc.add_paragraph(f"Votre offre a été retenue pour {f_amt} DHS TTC ({format_to_words_fr(f_amt)}).")
        
        doc.add_page_break()
        apply_official_header(doc, st.session_state['logo_data'])
        doc.add_heading(f"ORDRE DE SERVICE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"\nIl est prescrit à {f_soc} de commencer les travaux sous {f_delai} jours.")
        
        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 تحميل OS + Notification", bio.getvalue(), "OS_Notif.docx")

# --- التبويب الثالث: صور الأوراش ---
with tab_photos:
    st.subheader("📸 تقرير صور الأوراش")
    uploaded_imgs = st.file_uploader("ارفع صور الورش", accept_multiple_files=True)
    if st.button("🚀 توليد تقرير الصور"):
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        doc.add_heading("REPORTAGE PHOTOGRAPHIQUE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        for img in uploaded_imgs:
            doc.add_picture(BytesIO(img.getvalue()), width=Inches(4))
            doc.add_paragraph("الوصف: .........................")
        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 تحميل التقرير", bio.getvalue(), "Photos.docx")
