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
    st.info("Traitement et calcul complet en cours...")
    
    try:
        # Chargement du classeur avec openpyxl
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            max_row = ws_cmd.max_row
            
            # Application de la logique de calcul ligne par ligne
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    # Récupération sécurisée des données de la ligne
                    qte = df_cmd.loc[idx, 'Qté_commandée'] if 'Qté_commandée' in df_cmd.columns else 0
                    montant_st = df_cmd.loc[idx, 'Montant_ST'] if 'Montant_ST' in df_cmd.columns else 0
                    
                    qte_val = float(qte) if pd.notnull(qte) and str(qte).replace('.','',1).isdigit() else 0
                    st_val = float(montant_st) if pd.notnull(montant_st) and str(montant_st).replace('.','',1).isdigit() else 0
                    
                    rabais_total = qte_val * st_val
                    
                    # Injection directe des résultats calculés dans les cellules (A à V)
                    ws_cmd.cell(row=r, column=15).value = rabais_total # Colonne O: Rabais total
                    ws_cmd.cell(row=r, column=18).value = 10           # Colonne R: Tolérance par défaut
                    ws_cmd.cell(row=r, column=22).value = rabais_total # Colonne V: Écart
            
            # Sauvegarde en mémoire du fichier final propre (sans formules textuelles)
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées et calculées.")
            
            # 3. BOUTON DE TÉLÉCHARGEMENT
            st.download_button(
                label="📥 Télécharger le rapport final calculé (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final_Calcule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets 'Rabais fournisseurs' et/ou 'Rabais entre 2 dates' sont introuvables dans le fichier.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
