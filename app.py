import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. الدوال المساعدة (المبالغ والتنسيق) ---
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

def apply_official_header(doc, logo_data):
    section = doc.sections[0]
    header = section.header
    htable = header.add_table(1, 3, Inches(6.5))
    # Gauche (Français)
    c_left = htable.rows[0].cells[0].paragraphs[0]
    c_left.text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nPROVINCE DE TAROUDANT\nCOMMUNE D'ASKAOUN"
    c_left.style.font.size = Pt(9)
    # Centre (Logo)
    if logo_data:
        logo_run = htable.rows[0].cells[1].paragraphs[0].add_run()
        logo_run.add_picture(BytesIO(logo_data), width=Cm(1.8))
        htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Droite (Arabe)
    c_right = htable.rows[0].cells[2].paragraphs[0]
    c_right.text = "المملكة المغربية\nوزارة الداخلية\nإقليم تارودانت\nجماعة أسكاون"
    c_right.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    c_right.style.font.size = Pt(9)

# --- 2. إعدادات النظام ---
st.set_page_config(page_title="Système Intégré Askaouen", layout="wide")
if 'logo_data' not in st.session_state: st.session_state['logo_data'] = None

# --- 3. القائمة الجانبية (الشعار واللجنة) ---
st.sidebar.header("🖼️ Configuration Officielle")
up_logo = st.sidebar.file_uploader("Logo de la Commune", type=['png', 'jpg', 'jpeg'])
if up_logo: st.session_state['logo_data'] = up_logo.getvalue()

st.sidebar.divider()
st.sidebar.header("👥 Membres de la Commission")
all_members = [
    {"name": "MOHAMED ZILALI", "role": "Président de la commission"},
    {"name": "M BAREK BAK", "role": "Directeur du service"},
    {"name": "ATTAKY ABDELLATIF", "role": "Technicien de la commune"},
    {"name": "NOUREDDIN SALHI", "role": "Service des dépenses"},
    {"name": "FAYSSAL KADRI", "role": "Technicien à la commune"}
]
selected_members = [m for m in all_members if st.sidebar.checkbox(m['name'], value=True)]

# --- 4. واجهة التبويبات ---
tab_pv, tab_admin, tab_photos = st.tabs(["🏛️ المحاضر (PV 1-6)", "✉️ OS & Notification", "📸 تقارير الصور"])

# --- التبويب 1: المحاضر (استعادة القالب الأصلي بدقة) ---
with tab_pv:
    with st.expander("📝 Détails du Bon de Commande", expanded=True):
        c1, c2 = st.columns(2)
        num_bc = c1.text_input("N° BC", "01/ASK/2026")
        date_pub = c2.date_input("Date de publication", date(2026, 3, 1))
        obj_bc = st.text_area("Objet", "Location de matériel...")

    st.subheader("📊 Liste des concurrents")
    df_init = pd.DataFrame([{"Rang": i+1, "Nom": f"Société {i+1}", "Montant": "0.00"} for i in range(5)])
    data = st.data_editor(df_init, use_container_width=True)

    c3, c4, c5 = st.columns(3)
    pv_num = c3.selectbox("PV N°:", [1, 2, 3, 4, 5, 6])
    reu_date = c4.date_input("Date Séance", date.today())
    next_reu = c5.date_input("Prochain RDV", reu_date + timedelta(days=1))
    is_final = st.checkbox("✅ Attribution Finale (إسناد نهائي)")

    if st.button("🚀 Générer le PV Officiel"):
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        
        # العنوان
        pv_lbl = "1ére" if pv_num == 1 else f"{pv_num}éme"
        doc.add_heading(f"{pv_lbl} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph(f"Objet : {obj_bc}").bold = True
        doc.add_paragraph(f"Le {reu_date.strftime('%d/%m/%Y')} à 10h00, la commission d’ouverture des plis composée comme suit :")
        for m in selected_members: doc.add_paragraph(f"- {m['name']} : {m['role']}")

        doc.add_paragraph(f"\nS’est réunie dans la salle de la réunion de la commune sur invitation du président concernant l’avis d’achat n° {num_bc} publié le {date_pub.strftime('%d/%m/%Y')} sur le portail des marchés publics, ayant pour objet : {obj_bc}")

        idx = min(pv_num - 1, len(data)-1)
        curr = data.iloc[idx]
        amt_w = format_to_words_fr(curr['Montant'])

        if pv_num == 1:
            doc.add_paragraph("Les soumissionnaires qui ont déposés leurs offres électroniquement sont :")
            t = doc.add_table(rows=1, cols=3); t.style = 'Table Grid'
            for _, r in data.iterrows():
                row = t.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
            doc.add_paragraph(f"\nLe président invite {curr['Nom']} est le moins disant pour un montant de {curr['Montant']} Dhs TTC ({amt_w}) à confirmer son offre، suspend la séance et fixe un RDV le {next_reu.strftime('%d/%m/%Y')}.")
        
        elif is_final:
            doc.add_paragraph(f"Après vérification du portail... constate que la société : {curr['Nom']} a confirmé son offre.")
            p = doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le BC à {curr['Nom']} pour {curr['Montant']} Dhs TTC ({amt_w}).")
            p.bold = True
        
        else:
            prev = data.iloc[idx-1]
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {prev['Nom']} n’a pas confirmé son offre par lettre de confirmation.")
            doc.add_paragraph(f"Après écartement de la société {prev['Nom']} le président invite la société : {curr['Nom']} qui est classé le {pv_num}éme pour un montant de {curr['Montant']} Dhs TTC ({amt_w}) à confirmer son offre، suspend la séance et fixe un RDV le {next_reu.strftime('%d/%m/%Y')}.")

        # التوقيعات (3 سم فراغ)
        doc.add_paragraph(f"\nAskaouen le {reu_date.strftime('%d/%m/%Y')}")
        sig_t = doc.add_table(rows=0, cols=2); sig_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for i in range(0, len(selected_members), 2):
            r_n = sig_t.add_row(); r_n.cells[0].text = selected_members[i]['name']
            if i+1 < len(selected_members): r_n.cells[1].text = selected_members[i+1]['name']
            r_s = sig_t.add_row(); r_s.height = Cm(3.0)

        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 Télécharger PV", bio.getvalue(), f"PV_{pv_num}.docx")

# --- التبويب 2: OS & Notification (النماذج الرسمية) ---
with tab_admin:
    st.subheader("Ordre de Service & Notification")
    with st.form("admin_form"):
        soc_f = st.text_input("Société Adjudicataire")
        amt_f = st.text_input("Montant TTC")
        delai_f = st.number_input("Délai (jours)", 10)
        sub_admin = st.form_submit_button("🚀 Générer OS + Notification")
    
    if sub_admin:
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        doc.add_heading("ORDRE DE SERVICE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_paragraph(f"\nIl est prescrit à la société {soc_f} de commencer les travaux sous un délai de {delai_f} jours.")
        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 Télécharger Documents", bio.getvalue(), "Admin_Docs.docx")

# --- التبويب 3: تقرير الصور (Mise en page Photos) ---
with tab_photos:
    st.subheader("📸 Reportage Photos du Chantier")
    site_title = st.text_input("Titre du Projet")
    imgs = st.file_uploader("Upload Photos", accept_multiple_files=True)
    if st.button("🚀 Générer Rapport Photo"):
        doc = Document()
        apply_official_header(doc, st.session_state['logo_data'])
        doc.add_heading("REPORTAGE PHOTOGRAPHIQUE", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
        for img in imgs:
            doc.add_picture(BytesIO(img.getvalue()), width=Inches(4))
            doc.add_paragraph("Observation : .............................")
        bio = BytesIO(); doc.save(bio)
        st.download_button("📥 Télécharger Rapport", bio.getvalue(), "Photos.docx")
