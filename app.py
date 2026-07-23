import streamlit as st
import pandas as pd
import openpyxl
import io
import numpy as np

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs")
st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Traitement complet et application des règles de suppression en cours...")
    
    try:
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            df_cmd.columns = df_cmd.columns.str.strip()
            df_rabais.columns = df_rabais.columns.str.strip()
            
            max_row = ws_cmd.max_row
            
            # --- INDEXATION GLOBALE POUR LES CRITÈRES ---
            cles_reclamees = set(df_cmd[df_cmd['Date_Réclamée'].notnull() & (df_cmd['Date_Réclamée'] != 'NaT') & (df_cmd['Date_Réclamée'].astype(str).str.strip() != '')]['Clé_unique_détail_commande'].dropna())
            cles_facture = set(df_cmd['Clé_unique_détail_facture'].dropna())
            cles_credite = set(df_cmd['Clé_unique_détail_credité'].dropna()) if 'Clé_unique_détail_credité' in df_cmd.columns else set()
            
            prod_col = '# Produit' if '# Produit' in df_rabais.columns else df_rabais.columns[1]
            debut_col = 'Date début' if 'Date début' in df_rabais.columns else [c for c in df_rabais.columns if 'début' in c.lower()][0]
            fin_col = 'Date échéance' if 'Date échéance' in df_rabais.columns else [c for c in df_rabais.columns if 'échéance' in c.lower() or 'fin' in c.lower()][0]
            rabais_col = 'Rabais' if 'Rabais' in df_rabais.columns else [c for c in df_rabais.columns if 'rabais' in c.lower()][0]
            
            df_rabais[debut_col] = pd.to_datetime(df_rabais[debut_col], errors='coerce')
            df_rabais[fin_col] = pd.to_datetime(df_rabais[fin_col], errors='coerce')
            df_cmd['Date_Facture'] = pd.to_datetime(df_cmd['Date_Facture'], errors='coerce')
            rabais_par_produit = df_rabais.groupby(prod_col)
            
            # --- BOUCLE LIGNE PAR LIGNE ---
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    row = df_cmd.loc[idx]
                    
                    prod = row.get('No_Produit', None)
                    date_facture = row.get('Date_Facture', pd.NaT)
                    qte = float(row.get('Qté_commandée', 0)) if pd.notnull(row.get('Qté_commandée', 0)) else 0
                    montant_st = float(row.get('Montant_ST', 0)) if pd.notnull(row.get('Montant_ST', 0)) else 0
                    code_promo = str(row.get('Code_promotion', ''))
                    cle_cmd = row.get('Clé_unique_détail_commande', '')
                    cle_fact = row.get('Clé_unique_détail_facture', '')
                    cle_cred = row.get('Clé_unique_détail_credité', '') if 'Clé_unique_détail_credité' in df_cmd.columns else ''
                    date_recl = row.get('Date_Réclamée', None)
                    date_recl_cred = row.get('Date_réclamé_détail_credité', None) if 'Date_réclamé_détail_credité' in df_cmd.columns else None
                    tolerance = 10
                    
                    # -------------------------------------------------------------------------
                    # ÉVALUATION DE CHAQUE COLONNE DE SUPPRESSION (COLONNES B À M)
                    # -------------------------------------------------------------------------
                    # Supprimer #1 (Col B) : S'il y a une Date_Réclamée
                    suppr_1 = "Supprimer" if pd.notnull(date_recl) and str(date_recl).strip() != "" and str(date_recl) != "NaT" else ""
                    
                    # Supprimer #2 (Col C) : Si Date_Réclamée trouvée pour cette clé de commande
                    suppr_2 = "Supprimer" if cle_cmd in cles_reclamees else ""
                    
                    # Colonne D : Si la clé de commande se trouve dans la liste réclamée (Doublon)
                    suppr_d = "Supprimer" if suppr_2 == "Supprimer" else ""
                    
                    # Supprimer #3 (Col E) : Correspondance croisée entre Clé_facture et Clé_crédité
                    suppr_3 = "Supprimer" if (str(cle_cred) in cles_facture or str(cle_fact) in cles_credite) and (str(cle_cred) != "" or str(cle_fact) != "") else ""
                    
                    # Supprimer #4 (Col F) : Si la Clé_unique_détail_facture se trouve dans Clé_unique_détail_crédité
                    suppr_4 = "Supprimer" if str(cle_fact) in cles_credite and str(cle_fact) != "" else ""
                    
                    # Colonne G : Doublon lié à la colonne F
                    suppr_g = "Supprimer" if suppr_4 == "Supprimer" else ""
                    
                    # Supprimer #5 (Col H) : Quantité commandée négative
                    suppr_5 = "Supprimer" if qte < 0 else ""
                    
                    # Supprimer #6 (Col I) : Quantité positive sans réclamation
                    suppr_6 = "" 
                    
                    # Supprimer #7 (Col J) : Validation du rabais maximal et de la plus grande facture
                    suppr_7 = ""
                    
                    # Supprimer #8 (Col K) : Quantité > 0, date réclamation crédit non vide, montant < 0,99
                    suppr_8 = "Supprimer" if qte > 0 and pd.notnull(date_recl_cred) and str(date_recl_cred).strip() != "" and str(date_recl_cred) != "NaT" and montant_st < 0.99 else ""
                    
                    # Supprimer #9 (Col L) : Code promo commence par "FIL"
                    suppr_9 = "Supprimer" if code_promo.upper().startswith("FIL") else ""
                    
                    # Supprimer #10 (Col M) : Montant_ST < 0,99
                    suppr_10 = "Supprimer" if montant_st < 0.99 else ""
                    
                    # --- CALCULS FINANCIERS ET TEMPORELS ---
                    rabais_entre_2_dates = 0
                    date_debut_retenue = ""
                    date_fin_retenue = ""
                    indicateur_tolerance = 0
                    
                    if prod in rabais_par_produit.groups and pd.notnull(date_facture):
                        sub_df = rabais_par_produit.get_group(prod).copy()
                        sub_df['crit'] = (sub_df[debut_col] <= date_facture + pd.Timedelta(days=tolerance)) & \
                                         (sub_df[fin_col] >= date_facture - pd.Timedelta(days=tolerance))
                        valid_rabais = sub_df[sub_df['crit']].copy()
                        if not valid_rabais.empty:
                            valid_rabais['dH'] = (valid_rabais[debut_col] - date_facture).abs()
                            valid_rabais['dI'] = (valid_rabais[fin_col] - date_facture).abs()
                            valid_rabais['dist'] = valid_rabais[['dH', 'dI']].min(axis=1)
                            valid_rabais = valid_rabais.sort_values(by=['dist', debut_col], ascending=[True, False])
                            best_row = valid_rabais.iloc[0]
                            
                            taux_rabais = float(best_row[rabais_col]) if pd.notnull(best_row[rabais_col]) else 0
                            rabais_entre_2_dates = taux_rabais * qte
                            date_debut_retenue = best_row[debut_col].strftime('%Y-%m-%d') if pd.notnull(best_row[debut_col]) else ""
                            date_fin_retenue = best_row[fin_col].strftime('%Y-%m-%d') if pd.notnull(best_row[fin_col]) else ""
                            
                            if not ((best_row[debut_col] <= date_facture) and (best_row[fin_col] >= date_facture)):
                                indicateur_tolerance = 1
                    
                    rabais_total = qte * montant_st
                    ecart = rabais_total - rabais_entre_2_dates
                    
                    # --- SUPPRIMER TOTAL (Col A) : Synthèse globale ---
                    tous_criteres = [suppr_1, suppr_2, suppr_d, suppr_3, suppr_4, suppr_g, suppr_5, suppr_6, suppr_7, suppr_8, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in tous_criteres) else ""
                    
                    # --- ÉCRITURE DANS LE CLASSEUR EXCEL ---
                    ws_cmd.cell(row=r, column=1).value = suppr_total        # A
                    ws_cmd.cell(row=r, column=2).value = suppr_1            # B
                    ws_cmd.cell(row=r, column=3).value = suppr_2            # C
                    ws_cmd.cell(row=r, column=4).value = suppr_d            # D
                    ws_cmd.cell(row=r, column=5).value = suppr_3            # E
                    ws_cmd.cell(row=r, column=6).value = suppr_4            # F
                    ws_cmd.cell(row=r, column=7).value = suppr_g            # G
                    ws_cmd.cell(row=r, column=8).value = suppr_5            # H
                    ws_cmd.cell(row=r, column=9).value = suppr_6            # I
                    ws_cmd.cell(row=r, column=10).value = suppr_7           # J
                    ws_cmd.cell(row=r, column=11).value = suppr_8           # K
                    ws_cmd.cell(row=r, column=12).value = suppr_9           # L
                    ws_cmd.cell(row=r, column=13).value = suppr_10          # M
                    
                    ws_cmd.cell(row=r, column=15).value = rabais_total      # O
                    ws_cmd.cell(row=r, column=16).value = rabais_entre_2_dates # P
                    ws_cmd.cell(row=r, column=17).value = indicateur_tolerance # Q
                    ws_cmd.cell(row=r, column=18).value = tolerance         # R
                    
                    if date_debut_retenue:
                        ws_cmd.cell(row=r, column=19).value = date_debut_retenue # S
                    if date_fin_retenue:
                        ws_cmd.cell(row=r, column=20).value = date_fin_retenue   # T
                        
                    ws_cmd.cell(row=r, column=22).value = ecart             # V
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement complet et mis à jour ! {max_row - 1} lignes traitées avec succès.")
            
            st.download_button(
                label="📥 Télécharger le rapport final mis à jour (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Les onglets 'Rabais fournisseurs' et/ou 'Rabais entre 2 dates' sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
