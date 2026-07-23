import streamlit as st
import pandas as pd
import openpyxl
import io
import numpy as np

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs (Calcul Direct)")
st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Traitement et application des règles métier en cours...")
    
    try:
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            df_cmd.columns = df_cmd.columns.str.strip()
            df_rabais.columns = df_rabais.columns.str.strip()
            
            max_row = ws_cmd.max_row
            
            # Application des règles ligne par ligne
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    row_data = df_cmd.loc[idx]
                    
                    # Récupération des données nécessaires
                    date_reclamee = row_data.get('Date_Réclamée', None)
                    montant_st = float(row_data.get('Montant_ST', 0)) if pd.notnull(row_data.get('Montant_ST', 0)) else 0
                    qte = float(row_data.get('Qté_commandée', 0)) if pd.notnull(row_data.get('Qté_commandée', 0)) else 0
                    code_promo = str(row_data.get('Code_promotion', ''))
                    
                    # Règles de suppression (basées sur votre documentation détaillée)
                    # Supprimer #1 : S'il y a une Date_Réclamée
                    suppr_1 = "Supprimer" if pd.notnull(date_reclamee) and str(date_reclamee).strip() != "" and str(date_reclamee) != "NaT" else ""
                    
                    # Supprimer #9 : Si le Code de promotion commence par "FIL"
                    suppr_9 = "Supprimer" if code_promo.upper().startswith("FIL") else ""
                    
                    # Supprimer #10 : Si Montant_ST est inférieur à 0,99
                    suppr_10 = "Supprimer" if montant_st < 0.99 else ""
                    
                    # Rabais total = Qté * Montant_ST
                    rabais_total = qte * montant_st
                    
                    # Supprimer total : Si un critère de suppression est rencontré
                    critères_suppr = [suppr_1, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in critères_suppr) else ""
                    
                    # Inscription propre dans les cellules Excel correspondantes
                    ws_cmd.cell(row=r, column=1).value = suppr_total   # Supprimer total (A)
                    ws_cmd.cell(row=r, column=2).value = suppr_1       # Supprimer #1 (B)
                    ws_cmd.cell(row=r, column=12).value = suppr_9      # Supprimer #9 (L)
                    ws_cmd.cell(row=r, column=13).value = suppr_10     # Supprimer #10 (M)
                    
                    # Rabais total (O / 15)
                    ws_cmd.cell(row=r, column=15).value = rabais_total 
                    
                    # Tolérance (R / 18) et Indicateur de tolérance (Q / 17)
                    ws_cmd.cell(row=r, column=18).value = 10           
                    ws_cmd.cell(row=r, column=17).value = 0            
                    
                    # Écart Rabais total vs Entre 2 date (V / 22)
                    ws_cmd.cell(row=r, column=22).value = rabais_total 
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes analysées.")
            
            st.download_button(
                label="📥 Télécharger le rapport final conforme (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final_Calcule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets requis sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
