import streamlit as st
import pandas as pd
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, Cm
from io import BytesIO
from datetime import date, timedelta
from num2words import num2words

# --- 1. إعداد قاعدة بيانات الأرشيف ---
ARCHIVE_FILE = "log_pv_askaouen.csv"

def add_to_archive(num_bc, date_pv, obj, winner, amount):
    if not os.path.isfile(ARCHIVE_FILE):
        df = pd.DataFrame(columns=["Date_Generation", "N_BC", "Séance", "Objet", "Attributaire", "Montant"])
        df.to_csv(ARCHIVE_FILE, index=False, encoding='utf-16')
    
    new_entry = pd.DataFrame([[date.today(), num_bc, date_pv, obj, winner, amount]], 
                             columns=["Date_Generation", "N_BC", "Séance", "Objet", "Attributaire", "Montant"])
    new_entry.to_csv(ARCHIVE_FILE, mode='a', header=False, index=False, encoding='utf-16')

# --- 2. دالة المبالغ ---
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

st.set_page_config(page_title="Commune Askaouen - Archive & PV", layout="wide")

if 'logo_data' not in st.session_state: st.session_state['logo_data'] = None

# --- 3. الشريط الجانبي ---
st.sidebar.header("🖼️ Configuration & Logo")
uploaded_logo = st.sidebar.file_uploader("Logo officiel", type=['png', 'jpg'])
if uploaded_logo: st.session_state['logo_data'] = uploaded_logo.getvalue()

st.sidebar.divider()
st.sidebar.header("👥 Membres Présents")
all_members = [
    {"name": "MOHAMED ZILALI", "role": "Président de la commission"},
    {"name": "M BAREK BAK", "role": "Directeur du service"},
    {"name": "ATTAKY ABDELLATIF", "role": "Technicien de la commune"},
    {"name": "NOUREDDIN SALHI", "role": "Service des dépenses"},
    {"name": "FAYSSAL KADRI", "role": "Technicien à la commune"}
]
selected_members = [m for m in all_members if st.sidebar.checkbox(m['name'], value=True)]

# --- 4. التبويبات (Tabs) ---
tab_gen, tab_arch = st.tabs(["📝 Génération du PV", "📂 Archive des Bons"])

with tab_gen:
    st.title("🏛️ Système de Gestion des PV - Askaouen")
    
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
    next_meeting = st.date_input("Prochain RDV", reunion_date + timedelta(days=1))
    is_final = st.checkbox("✅ Attribution Finale (إسناد)")

    if st.button("🚀 Générer & Archiver"):
        if not selected_members:
            st.error("Veuillez sélectionner au moins un membre!")
        else:
            doc = Document()
            section = doc.sections[0]
            section.left_margin, section.right_margin = Cm(2.5), Cm(2)

            # --- Header ---
            header = section.header
            htable = header.add_table(1, 3, Inches(6.5))
            htable.rows[0].cells[0].paragraphs[0].text = "ROYAUME DU MAROC\nMINISTERE DE L'INTERIEUR\nCOMMUNE D'ASKAOUN"
            if st.session_state['logo_data']:
                logo_run = htable.rows[0].cells[1].paragraphs[0].add_run()
                logo_run.add_picture(BytesIO(st.session_state['logo_data']), width=Cm(1.8))
                htable.rows[0].cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            htable.rows[0].cells[2].paragraphs[0].text = "المملكة المغربية\nوزارة الداخلية\nجماعة أسكاون"
            htable.rows[0].cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

            # --- Body ---
            pv_lbl = "1ére" if pv_num == 1 else f"{pv_num}éme"
            doc.add_heading(f"{pv_lbl} Procès verbal", 1).alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("De la commission d’ouverture des plis\nProcédure Bon de commande").alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph(f"Objet : {obj_bc}").bold = True
            doc.add_paragraph(f"Le {reunion_date.strftime('%d/%m/%Y')} à {reunion_hour}, la commission d’ouverture des plis composée comme suit :")
            for m in selected_members:
                doc.add_paragraph(f"- {m['name']} : {m['role']}")

            doc.add_paragraph(f"\nS’est réunie dans la salle de la réunion de la commune sur invitation du président concernant l’avis d’achat du bon de commande n° {num_bc} publié le : {date_pub.strftime('%d/%m/%Y')} sur le portail des marchés publics, ayant pour objet : {obj_bc}")

            idx = min(pv_num - 1, len(data)-1)
            curr_co = data.iloc[idx]
            amt_w = format_to_words_fr(curr_co['Montant'])

            if pv_num == 1:
                doc.add_paragraph("Les soumissionnaires qui ont déposés leurs offres électroniquement sont :")
                tab = doc.add_table(rows=1, cols=3); tab.style = 'Table Grid'
                for _, r in data.iterrows():
                    row = tab.add_row().cells
                    row[0].text, row[1].text, row[2].text = str(r['Rang']), r['Nom'], f"{r['Montant']} MAD"
                doc.add_paragraph(f"\nFormat papier : Néant.\nLe président invite la société : {curr_co['Nom']} est le moins disant pour {curr_co['Montant']} Dhs TTC ({amt_w}) à confirmer son offre، suspend la séance et fixe un RDV le {next_meeting.strftime('%d/%m/%Y')}.")
            
            elif is_final:
                doc.add_paragraph(f"Après vérification... la société : {curr_co['Nom']} a confirmé son offre.")
                doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le BC à {curr_co['Nom']} لبلغ {curr_co['Montant']} Dhs TTC {amt_w}.").bold = True
                # تسجيل في الأرشيف
                add_to_archive(num_bc, reunion_date, obj_bc, curr_co['Nom'], curr_co['Montant'])
            
            else:
                prev_co = data.iloc[idx-1]
                doc.add_paragraph(f"La société {prev_co['Nom']} n’a pas confirmé. Le président invite {curr_co['Nom']} (classé {pv_num}éme) بمبلغ {curr_co['Montant']} Dhs ({amt_w}) à confirmer، RDV le {next_meeting.strftime('%d/%m/%Y')}.")

            # --- Signatures ---
            doc.add_paragraph(f"\nAskaouen le {reunion_date.strftime('%d/%m/%Y')}\n")
            sig_tab = doc.add_table(rows=0, cols=2)
            sig_tab.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for i in range(0, len(selected_members), 2):
                r_n = sig_tab.add_row()
                r_n.cells[0].text = selected_members[i]['name']
                r_n.cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                if i+1 < len(selected_members):
                    r_n.cells[1].text = selected_members[i+1]['name']
                    r_n.cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_s = sig_tab.add_row(); r_s.height = Cm(2.5)

            bio = BytesIO()
            doc.save(bio)
            st.success("✅ PV Généré et données sauvegardées !")
            st.download_button("📥 Télécharger le Document", bio.getvalue(), f"PV_{num_bc.replace('/','_')}.docx")

with tab_arch:
    st.header("📂 Archive des Bons de Commande Attribués")
    if os.path.isfile(ARCHIVE_FILE):
        archive_df = pd.read_csv(ARCHIVE_FILE, encoding='utf-16')
        st.dataframe(archive_df, use_container_width=True)
        st.download_button("📥 Exporter l'archive (Excel/CSV)", archive_df.to_csv(index=False).encode('utf-16'), "Archive_Askaouen.csv")
    else:
        st.info("L'archive هو فارغ حالياً.")
