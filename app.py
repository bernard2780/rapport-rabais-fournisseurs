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
    st.info("Traitement et application complète des règles en cours...")
    
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
            rabais_col = 'Rabais' if 'Rabais' in df_rabais.columns else [c for c in df_rabais.columns if 'rabais' in c.lower()][0]
            
            rabais_dict = df_rabais.groupby(prod_col_rabais)[rabais_col].max().to_dict()
            
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    no_produit = df_cmd.loc[idx, 'No_Produit'] if 'No_Produit' in df_cmd.columns else None
                    qte = df_cmd.loc[idx, 'Qté_commandée'] if 'Qté_commandée' in df_cmd.columns else 0
                    
                    qte_val = float(qte) if pd.notnull(qte) and str(qte).replace('.','',1).isdigit() else 0
                    
                    montant_rabais_unitaire = rabais_dict.get(no_produit, 0)
                    if pd.isnull(montant_rabais_unitaire):
                        montant_rabais_unitaire = 0
                        
                    rabais_total_calc = qte_val * float(montant_rabais_unitaire)
                    
                    # Condition de suppression / marquage
                    do_suppress = (no_produit not in rabais_dict) or (qte_val <= 0) or (rabais_total_calc <= 0)
                    
                    valeur_suppr = "Supprimer" if do_suppress else ""
                    
                    # Remplissage de TOUTES les colonnes de suppression (de la colonne 1 à 13)
                    for col_idx in range(1, 14):
                        ws_cmd.cell(row=r, column=col_idx).value = valeur_suppr
                    
                    # Colonne 14: Rabais maximum offert
                    ws_cmd.cell(row=r, column=14).value = float(montant_rabais_unitaire)
                    
                    # Colonne 15: Rabais total
                    ws_cmd.cell(row=r, column=15).value = rabais_total_calc 
                    
                    # Colonne 18: Indicateur de tolérance
                    ws_cmd.cell(row=r, column=18).value = 10                  
                    
                    # Colonne 22: Écart
                    ws_cmd.cell(row=r, column=22).value = rabais_total_calc   
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement complet terminé ! {max_row - 1} lignes traitées.")
            
            st.download_button(
                label="📥 Télécharger le rapport final complet (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final_Calcule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets requis sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
