# --- الجزء الخاص بمنطق محتوى المحضر (داخل زر الإنشاء) ---

        # 1. حالة المحضر الأول (دائماً استدعاء الشركة الأولى)
        if pv_num == 1:
            doc.add_paragraph("Après vérification du portail des marchés publics, les soumissionnaires qui ont déposé leurs offres de prix électroniquement sont :")
            tab = doc.add_table(rows=1, cols=3)
            tab.style = 'Table Grid'
            hdr = tab.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Rang', 'Concurrent', 'Montant TTC'
            for _, r in data.iterrows():
                row = tab.add_row().cells
                row[0].text, row[1].text, row[2].text = str(r['Rang']), str(r['Nom']), f"{r['Montant']} MAD"

            first_co = data.iloc[0]
            amt_w = format_to_words_fr(first_co['Montant'])
            doc.add_paragraph(f"\nAprès examen des offres, le président de la commission invite la société : {first_co['Nom']} (Moins disant) pour un montant de {first_co['Montant']} Dhs TTC ({amt_w}) à confirmer son offre par lettre de confirmation.")
            
            # هنا يمكنك إضافة الجملة التي تنقصك في المحضر الأول
            doc.add_paragraph("\n[أضف الجملة الناقصة هنا]")

        # 2. حالة "المحضر النهائي" (سواء كان 2 أو 3 أو 6 أو غيره)
        elif is_final_attr:
            current_idx = min(pv_num - 1, len(data) - 1)
            curr_co = data.iloc[current_idx]
            amt_w = format_to_words_fr(curr_co['Montant'])
            
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {curr_co['Nom']} a confirmé son offre par lettre de confirmation.")
            p_res = doc.add_paragraph(f"Le président VALIDE la confirmation et ATTRIBUE le bon de commande à la société {curr_co['Nom']} pour un montant de : {curr_co['Montant']} Dhs TTC ({amt_w}).")
            p_res.bold = True
            
            # الجملة الختامية التي تظهر في المحاضر النهائية
            doc.add_paragraph("\nLa séance a été levée à [وقت انتهاء الجلسة].")

        # 3. حالة "غير مثمر" (Infructueux)
        elif is_infructueux:
            current_idx = min(pv_num - 1, len(data) - 1)
            curr_co = data.iloc[current_idx]
            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {curr_co['Nom']} n’a pas confirmé son offre.")
            p_inf = doc.add_paragraph("\nPAR CONSEQUENT, LA COMMISSION DECLARE QUE CE BON DE COMMANDE EST :")
            p_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            res_inf = doc.add_paragraph("INFRUCTUEUX")
            res_inf.alignment = WD_ALIGN_PARAGRAPH.CENTER
            res_inf.bold = True

        # 4. حالة المحاضر الانتقالية (رفض شركة واستدعاء التالية)
        else:
            current_idx = min(pv_num - 1, len(data) - 1)
            prev_idx = max(0, current_idx - 1)
            curr_co = data.iloc[current_idx]
            prev_co = data.iloc[prev_idx]
            amt_w = format_to_words_fr(curr_co['Montant'])

            doc.add_paragraph(f"Après vérification du portail des marchés publics, la commission constate que la société {prev_co['Nom']} n’a pas confirmé son offre par lettre de confirmation.")
            doc.add_paragraph(f"Après écartement de la société {prev_co['Nom']}, le président de la commission invite la société : {curr_co['Nom']} الذي يحتل الرتبة {pv_num} بمبلغ {curr_co['Montant']} Dhs TTC ({amt_w}) لتأكيد عرضه.")
