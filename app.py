import streamlit as st
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. الدوال المساعدة للنصوص الرسمية ---
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
    # يسار
    c_left = htable.rows[0].cells[0].paragraphs[0]
    c_left.text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nPROVINCE DE TAROUDANT\nCOMMUNE D'ASKAOUN"
    c_left.style.font.size = Pt(9)
    # وسط
    if logo_data:
        logo_run = htable.rows[0].cells[1].paragraphs[0].add_run()
        logo_run.add_picture(BytesIO(logo_data), width=Cm(1.8))
        htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # يمين
    c_right = htable.rows[0].cells[2].paragraphs[0]
    c_right.text = "المملكة المغربية\nوزارة الداخلية\nإقليم تارودانت\nجماعة أسكاون"
    c_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c_right.style.font.size = Pt(9)

st.set_page_config(page_title="Système Intégré Askaouen", layout="wide")

# --- 2. تبويبات النظام ---
tab_pv, tab_os_notif, tab_reception, tab_photos = st.tabs([
    "🏛️ محاضر فتح الأظرفة", "✉️ الإشعار و OS", "🏗️ تسليم الأشغال", "📸 صور الأوراش"
])

# --- التبويب 2: النماذج الرسمية لـ OS و Notification ---
with tab_os_notif:
    st.subheader("Génération de la Notification et l'Ordre de Service")
    with st.form("os_form"):
        col1, col2 = st.columns(2)
        with col1:
            st_societe = st.text_input("Nom de la Société (الفائزة)", "STE ...")
            st_montant = st.text_input("Montant de l'offre (DHS TTC)", "0.00")
            st_objet = st.text_area("Objet du BC (موضوع الصفقة)", "Location de matériel...")
        with col2:
            st_bc_num = st.text_input("N° Bon de Commande", "01/ASK/2026")
            st_os_num = st.text_input("N° Ordre de Service", "01/2026")
            st_date_start = st.date_input("Date de commencement (تاريخ البدء)", date.today())
            st_delai = st.number_input("Délai d'exécution (بالأيام)", value=10)
        
        submit_os = st.form_submit_button("🚀 توليد الوثائق الرسمية")

    if submit_os:
        doc = Document()
        
        # --- النموذج 1: Lettre de Notification (رسمي) ---
        apply_official_header(doc, st.session_state.get('logo_data'))
        doc.add_paragraph(f"\nAskaouen, le {date.today().strftime('%d/%m/%Y')}")
        notif_title = doc.add_heading("LETTRE DE NOTIFICATION", 1)
        notif_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"\nA Monsieur le Gérant de la société : {st_societe}").bold = True
        doc.add_paragraph(f"\nObjet : Notification de l'approbation du Bon de Commande n° {st_bc_num}")
        doc.add_paragraph(f"Référence : Avis d'achat n° {st_bc_num}")
        
        p1 = doc.add_paragraph(f"\nJ'ai l'honneur de vous informer que votre offre concernant l'objet cité ci-dessus a été retenue par la commission d'ouverture des plis pour un montant de :")
        doc.add_paragraph(f"{st_montant} DHS TTC ({format_to_words_fr(st_montant)}).").bold = True
        
        doc.add_paragraph("\nA cet effet, je vous invite à prendre contact avec les services de la commune pour la suite de la procédure.")
        doc.add_paragraph("\nVeuillez agréer, Monsieur le Gérant, l'expression de mes salutations distinguées.")
        
        doc.add_page_break()

        # --- النموذج 2: Ordre de Service (رسمي) ---
        apply_official_header(doc, st.session_state.get('logo_data'))
        doc.add_paragraph(f"\nAskaouen, le {date.today().strftime('%d/%m/%Y')}")
        os_title = doc.add_heading(f"ORDRE DE SERVICE N° {st_os_num}", 1)
        os_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"\nMarché/BC n° : {st_bc_num}")
        doc.add_paragraph(f"Objet : {st_objet}")
        doc.add_paragraph(f"A l'entreprise : {st_societe}").bold = True
        
        doc.add_paragraph(f"\nIl vous est prescrit par la présente de commencer l'exécution des prestations/travaux faisant l'objet du bon de commande cité ci-dessus à compter du :")
        doc.add_paragraph(f"{st_date_start.strftime('%d/%m/%Y')}").bold = True
        
        doc.add_paragraph(f"\nLe délai d'exécution qui vous est imparti est de : {st_delai} jours.")
        
        # مكان التوقيع
        sig_os = doc.add_table(1, 2, width=Inches(6))
        sig_os.cell(0, 0).text = "Accusé de réception par l'entreprise\n(Date et Signature)"
        sig_os.cell(0, 1).text = "Le Président de la Commune\n(Signature et Cachet)"
        sig_os.rows[0].height = Cm(3)

        bio = BytesIO(); doc.save(bio)
        st.success("✅ تم توليد الإشعار وأمر الخدمة بنجاح!")
        st.download_button("📥 تحميل المستندات (OS + Notif)", bio.getvalue(), f"OS_Notification_{st_societe}.docx")

# --- التبويب 3: محضر تسليم الأشغال (رسمي) ---
with tab_reception:
    st.subheader("Procès-Verbal de Réception des Travaux")
    with st.form("recep_form"):
        r_societe = st.text_input("Société", "STE ...")
        r_bc = st.text_input("N° BC", "01/ASK/2026")
        r_date = st.date_input("Date de Réception", date.today())
        r_type = st.selectbox("Type de réception", ["Provisoire (مؤقت)", "Définitive (نهائي)"])
        r_obs = st.text_area("Observations", "Conforme aux prescriptions techniques.")
        submit_r = st.form_submit_button("🚀 توليد محضر التسليم الرسمي")

    if submit_r:
        doc = Document()
        apply_official_header(doc, st.session_state.get('logo_data'))
        
        title_r = doc.add_heading(f"PROCES-VERBAL DE RECEPTION {r_type.upper()}", 1)
        title_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_paragraph(f"\nLe {r_date.strftime('%d/%m/%Y')}, la commission composée de :")
        for m in selected_members:
            doc.add_paragraph(f"- {m['name']} : {m['role']}")
            
        doc.add_paragraph(f"\nS'est rendue sur les lieux pour procéder à la réception des travaux faisant l'objet du BC n° {r_bc} exécuté par la société {r_societe}.")
        doc.add_paragraph(f"\nAprès examen et vérification, la commission constate que les prestations sont : {r_obs}")
        doc.add_paragraph("\nEn conséquence, la commission déclare la réception des travaux sus-indiqués.")
        
        # جدول توقيعات اللجنة
        doc.add_paragraph("\nSignatures des membres de la commission :")
        sig_tab = doc.add_table(rows=0, cols=2)
        for i in range(0, len(selected_members), 2):
            row = sig_tab.add_row()
            row.cells[0].text = selected_members[i]['name']
            if i+1 < len(selected_members):
                row.cells[1].text = selected_members[i+1]['name']
            r_sp = sig_tab.add_row(); r_sp.height = Cm(2.5)

        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 تحميل محضر التسليم", bio.getvalue(), "PV_Reception.docx")
