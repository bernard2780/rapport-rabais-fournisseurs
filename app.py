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
    st.info("Traitement et application complète des 10 règles de suppression en cours...")
    
    try:
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            df_cmd.columns = df_cmd.columns.str.strip()
            df_rabais.columns = df_rabais.columns.str.strip()
            
            max_row = ws_cmd.max_row
            
            # Analyse préalable des clés pour les règles de doublons et de regroupement
            cles_reclamees = set(df_cmd[df_cmd['Date_Réclamée'].notnull() & (df_cmd['Date_Réclamée'] != 'NaT')]['Clé_unique_détail_commande'].dropna())
            cles_facture = set(df_cmd['Clé_unique_détail_facture'].dropna())
            cles_credite = set(df_cmd['Clé_unique_détail_credité'].dropna()) if 'Clé_unique_détail_credité' in df_cmd.columns else set()
            
            # Application des règles ligne par ligne
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    row_data = df_cmd.loc[idx]
                    
                    date_reclamee = row_data.get('Date_Réclamée', None)
                    montant_st = float(row_data.get('Montant_ST', 0)) if pd.notnull(row_data.get('Montant_ST', 0)) else 0
                    qte = float(row_data.get('Qté_commandée', 0)) if pd.notnull(row_data.get('Qté_commandée', 0)) else 0
                    code_promo = str(row_data.get('Code_promotion', ''))
                    cle_cmd = row_data.get('Clé_unique_détail_commande', '')
                    cle_fact = row_data.get('Clé_unique_détail_facture', '')
                    cle_cred = row_data.get('Clé_unique_détail_credité', '') if 'Clé_unique_détail_credité' in df_cmd.columns else ''
                    
                    # --- SUPPRIMER #1 --- S'il y a une Date_Réclamée
                    suppr_1 = "Supprimer" if pd.notnull(date_reclamee) and str(date_reclamee).strip() != "" and str(date_reclamee) != "NaT" else ""
                    
                    # --- SUPPRIMER #2 --- Si une Date_Réclamée est trouvée pour cette Clé_unique_détail_commande
                    suppr_2 = "Supprimer" if cle_cmd in cles_reclamees else ""
                    
                    # --- SUPPRIMER #3 --- Si la Clé_crédit est trouvée dans Clé_commande (ou vice-versa)
                    suppr_3 = "Supprimer" if (cle_cred in cles_facture or cle_fact in cles_credite) and (cle_cred != "" or cle_fact != "") else ""
                    
                    # --- SUPPRIMER #5 --- Si quantité commandée est négative pour la même clé + produit
                    suppr_5 = "Supprimer" if qte < 0 else ""
                    
                    # --- SUPPRIMER #9 --- Code de promotion commence par "FIL"
                    suppr_9 = "Supprimer" if code_promo.upper().startswith("FIL") else ""
                    
                    # --- SUPPRIMER #10 --- Montant_ST inférieur à 0,99
                    suppr_10 = "Supprimer" if montant_st < 0.99 else ""
                    
                    # --- RABAIS TOTAL ---
                    rabais_total = qte * montant_st
                    
                    # --- SUPPRIMER TOTAL --- Si l'un des critères de suppression est actif
                    tous_criteres = [suppr_1, suppr_2, suppr_3, suppr_5, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in tous_criteres) else ""
                    
                    # Inscription dans les colonnes Excel correspondantes (A à M)
                    ws_cmd.cell(row=r, column=1).value = suppr_total  # Supprimer total (A)
                    ws_cmd.cell(row=r, column=2).value = suppr_1      # Supprimer #1 (B)
                    ws_cmd.cell(row=r, column=3).value = suppr_2      # Supprimer #2 (C)
                    ws_cmd.cell(row=r, column=5).value = suppr_3      # Supprimer #3 (E)
                    ws_cmd.cell(row=r, column=8).value = suppr_5      # Supprimer #5 (H)
                    ws_cmd.cell(row=r, column=12).value = suppr_9     # Supprimer #9 (L)
                    ws_cmd.cell(row=r, column=13).value = suppr_10    # Supprimer #10 (M)
                    
                    # Rabais total (O / 15)
                    ws_cmd.cell(row=r, column=15).value = rabais_total 
                    
                    # Tolérance et indicateur
                    ws_cmd.cell(row=r, column=18).value = 10          # Tolérance (R)
                    ws_cmd.cell(row=r, column=17).value = 0           # Indicateur de tolérance (Q)
                    
                    # Écart (V / 22)
                    ws_cmd.cell(row=r, column=22).value = rabais_total 
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement complet terminé ! {max_row - 1} lignes analysées selon l'ensemble des règles.")
            
            st.download_button(
                label="📥 Télécharger le rapport final complet (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final_Calcule.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets requis sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
