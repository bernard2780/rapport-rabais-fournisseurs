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
    st.info("Traitement et application des règles de suppression et de rabais en cours...")
    
    try:
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            df_cmd.columns = df_cmd.columns.str.strip()
            df_rabais.columns = df_rabais.columns.str.strip()
            
            max_row = ws_cmd.max_row
            
            prod_col_rabais = '# Produit' if '# Produit' in df_rabais.columns else df_rabais.columns[1]
            rabais_dict = df_rabais.set_index(prod_col_rabais).to_dict(orient='index')
            
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    no_produit = df_cmd.loc[idx, 'No_Produit'] if 'No_Produit' in df_cmd.columns else None
                    qte = df_cmd.loc[idx, 'Qté_commandée'] if 'Qté_commandée' in df_cmd.columns else 0
                    
                    qte_val = float(qte) if pd.notnull(qte) and str(qte).replace('.','',1).isdigit() else 0
                    
                    montant_rabais_unitaire = 0
                    produit_trouve = False
                    
                    if no_produit in rabais_dict:
                        produit_trouve = True
                        for col_name in ['Rabais', 'Montant_rabais', 'Taux']:
                            if col_name in rabais_dict[no_produit]:
                                val = rabais_dict[no_produit][col_name]
                                if pd.notnull(val) and str(val).replace('.','',1).isdigit():
                                    montant_rabais_unitaire = float(val)
                                    break
                    
                    rabais_total_calc = qte_val * montant_rabais_unitaire
                    
                    # Application des règles de marquage (Similaire à vos conditions de suppression)
                    # Si le produit n'a pas de rabais valide ou quantité nulle, on marque pour suppression
                    if not produit_trouve or qte_val <= 0 or rabais_total_calc <= 0:
                        ws_cmd.cell(row=r, column=1).value = "Supprimer"  # Colonne A: Supprimer total
                        ws_cmd.cell(row=r, column=2).value = "Supprimer"  # Colonne B: Supprimer #1
                    else:
                        ws_cmd.cell(row=r, column=1).value = ""
                        ws_cmd.cell(row=r, column=2).value = ""

                    # Inscription des valeurs calculées
                    ws_cmd.cell(row=r, column=15).value = rabais_total_calc # Colonne O: Rabais total
                    ws_cmd.cell(row=r, column=18).value = 10                  # Colonne R: Tolérance
                    ws_cmd.cell(row=r, column=22).value = rabais_total_calc   # Colonne V: Écart
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement et filtrage terminés ! {max_row - 1} lignes analysées.")
            
            st.download_button(
                label="📥 Télécharger le rapport filtré et calculé (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final_Calcule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets requis sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
