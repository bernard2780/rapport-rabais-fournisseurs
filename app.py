import streamlit as st
import pandas as pd
import openpyxl
import io
import numpy as np

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")

if 'version_compteur' not in st.session_state:
    st.session_state.version_compteur = 1

st.title(f"Générateur de Rapport : Rabais Fournisseurs (v{st.session_state.version_compteur})")
st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info(f"Traitement complet (Version {st.session_state.version_compteur}) en cours...")
    
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
            cles_facture = set(df_cmd['Clé_unique_détail_facture'].dropna().astype(str)) if 'Clé_unique_détail_facture' in df_cmd.columns else set()
            
            col_cred = 'Clé_unique_détail_credité' if 'Clé_unique_détail_credité' in df_cmd.columns else ([c for c in df_cmd.columns if 'crédit' in str(c).lower()][0] if any('crédit' in str(c).lower() for c in df_cmd.columns) else None)
            cles_credite = set(df_cmd[col_cred].dropna().astype(str)) if col_cred else set()
            
            prod_col = '# Produit' if '# Produit' in df_rabais.columns else df_rabais.columns[1]
            debut_col = 'Date début' if 'Date début' in df_rabais.columns else [c for c in df_rabais.columns if 'début' in str(c).lower()][0]
            fin_col = 'Date échéance' if 'Date échéance' in df_rabais.columns else [c for c in df_rabais.columns if 'échéance' in str(c).lower() or 'fin' in str(c).lower()][0]
            rabais_col = 'Rabais' if 'Rabais' in df_rabais.columns else [c for c in df_rabais.columns if 'rabais' in str(c).lower()][0]
            
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
                    
                    # Sécurisation stricte du code promo contre les valeurs non textuelles (float/NaN)
                    raw_promo = row.get('Code_promotion', '')
                    code_promo = str(raw_promo) if pd.notnull(raw_promo) and raw_promo != 'nan' else ''
                    
                    cle_cmd = str(row.get('Clé_unique_détail_commande', ''))
                    cle_fact = str(row.get('Clé_unique_détail_facture', '')) if pd.notnull(row.get('Clé_unique_détail_facture', '')) else ''
                    cle_cred_val = str(row.get(col_cred, '')) if col_cred and pd.notnull(row.get(col_cred, '')) else ''
                    date_recl = row.get('Date_Réclamée', None)
                    date_recl_cred = row.get('Date_réclamé_détail_credité', None) if 'Date_réclamé_détail_credité' in df_cmd.columns else None
                    tolerance = 10
                    
                    # --- ÉVALUATION DE CHAQUE COLONNE DE SUPPRESSION (B À M) ---
                    suppr_1 = "Supprimer" if pd.notnull(date_recl) and str(date_recl).strip() != "" and str(date_recl) != "NaT" else ""
                    suppr_2 = "Supprimer" if cle_cmd in cles_reclamees else ""
                    suppr_d = "Supprimer" if suppr_2 == "Supprimer" else ""
                    suppr_3 = "Supprimer" if ((cle_cred_val in cles_facture and cle_cred_val != '' and cle_cred_val != 'nan') or 
                                               (cle_fact in cles_credite and cle_fact != '' and cle_fact != 'nan')) else ""
                    suppr_4 = "Supprimer" if (cle_fact in cles_credite and cle_fact != '' and cle_fact != 'nan') else ""
                    suppr_g = "Supprimer" if suppr_4 == "Supprimer" else ""
                    suppr_5 = "Supprimer" if qte < 0 else ""
                    suppr_6 = "" 
                    suppr_7 = ""
                    has_date_recl_cred = pd.notnull(date_recl_cred) and str(date_recl_cred).strip() != "" and str(date_recl_cred) != "NaT"
                    suppr_8 = "Supprimer" if (qte > 0 and has_date_recl_cred and (cle_cred_val == '' or cle_cred_val == 'nan' or cle_cred_val == '0') and montant_st < 0.99) else ""
                    suppr_9 = "Supprimer" if code_promo.upper().startswith("FIL") else ""
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
                    
                    # --- SUPPRIMER TOTAL (Col A) ---
                    tous_criteres = [suppr_1, suppr_2, suppr_d, suppr_3, suppr_4, suppr_g, suppr_5, suppr_6, suppr_7, suppr_8, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in tous_criteres) else ""
                    
                    # --- ÉCRITURE DANS LE CLASSEUR EXCEL ---
                    ws_cmd.cell(row=r, column=1).value = suppr_total
                    ws_cmd.cell(row=r, column=2).value = suppr_1
                    ws_cmd.cell(row=r, column=3).value = suppr_2
                    ws_cmd.cell(row=r, column=4).value = suppr_d
                    ws_cmd.cell(row=r, column=5).value = suppr_3
                    ws_cmd.cell(row=r, column=6).value = suppr_4
                    ws_cmd.cell(row=r, column=7).value = suppr_g
                    ws_cmd.cell(row=r, column=8).value = suppr_5
                    ws_cmd.cell(row=r, column=9).value = suppr_6
                    ws_cmd.cell(row=r, column=10).value = suppr_7
                    ws_cmd.cell(row=r, column=11).value = suppr_8
                    ws_cmd.cell(row=r, column=12).value = suppr_9
                    ws_cmd.cell(row=r, column=13).value = suppr_10
                    
                    ws_cmd.cell(row=r, column=15).value = rabais_total
                    ws_cmd.cell(row=r, column=16).value = rabais_entre_2_dates
                    ws_cmd.cell(row=r, column=17).value = indicateur_tolerance
                    ws_cmd.cell(row=r, column=18).value = tolerance
                    
                    if date_debut_retenue:
                        ws_cmd.cell(row=r, column=19).value = date_debut_retenue
                    if date_fin_retenue:
                        ws_cmd.cell(row=r, column=20).value = date_fin_retenue
                        
                    ws_cmd.cell(row=r, column=22).value = ecart
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            nom_fichier = f"Rapport_Rabais_Final_v{st.session_state.version_compteur}.xlsx"
            
            st.success(f"Traitement terminé avec succès ! ({max_row - 1} lignes traitées)")
            
            if st.download_button(
                label=f"📥 Télécharger le rapport ({nom_fichier})",
                data=output_buffer,
                file_name=nom_fichier,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ):
                st.session_state.version_compteur += 1
        else:
            st.error("Les onglets 'Rabais fournisseurs' et/ou 'Rabais entre 2 dates' sont introuvables.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
