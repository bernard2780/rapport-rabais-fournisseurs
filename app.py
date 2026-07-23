import streamlit as st
import pandas as pd
import openpyxl
import io
import numpy as np
from datetime import datetime
from zoneinfo import ZoneInfo

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")

if 'version_compteur' not in st.session_state:
    st.session_state.version_compteur = 1

st.title(f"Générateur de Rapport : Rabais Fournisseurs (v{st.session_state.version_compteur})")

montreal_time = datetime.now(ZoneInfo("America/Montreal")).strftime("%Y-%m-%d à %H:%M:%S")
st.caption(f"📅 Génération du rapport — Heure de Montréal : {montreal_time}")

st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info(f"Traitement rigoureux (Version {st.session_state.version_compteur}) en cours...")
    
    try:
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames and 'Rabais entre 2 dates' in wb.sheetnames:
            ws_cmd = wb['Rabais fournisseurs']
            
            df_cmd = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
            df_rabais = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
            
            df_cmd.columns = df_cmd.columns.str.strip()
            df_rabais.columns = df_rabais.columns.str.strip()
            
            max_row = ws_cmd.max_row
            tolerance = 10
            
            # Nettoyage des types numériques de base
            df_cmd['Qté_commandée'] = pd.to_numeric(df_cmd['Qté_commandée'], errors='coerce').fillna(0)
            df_cmd['Montant_ST'] = pd.to_numeric(df_cmd['Montant_ST'], errors='coerce').fillna(0)
            df_cmd['Date_Facture'] = pd.to_datetime(df_cmd['Date_Facture'], errors='coerce')
            
            # --- IDENTIFICATION DES COLONNES CLÉS ---
            col_cred = 'Clé_unique_détail_credité' if 'Clé_unique_détail_credité' in df_cmd.columns else ([c for c in df_cmd.columns if 'crédit' in str(c).lower() and 'clé' in str(c).lower()][0] if any('crédit' in str(c).lower() and 'clé' in str(c).lower() for c in df_cmd.columns) else None)
            col_date_recl_cred = 'Date_réclamé_détail_credité' if 'Date_réclamé_détail_credité' in df_cmd.columns else ([c for c in df_cmd.columns if 'crédit' in str(c).lower() and 'date' in str(c).lower()][0] if any('crédit' in str(c).lower() and 'date' in str(c).lower() for c in df_cmd.columns) else None)
            
            col_promo = None
            for candidate in ['Code_de_Promotion', 'Code_promotion', 'Code promotion', 'Code de Promotion']:
                if candidate in df_cmd.columns:
                    col_promo = candidate
                    break
            if not col_promo:
                col_promo = next((c for c in df_cmd.columns if 'promo' in str(c).lower()), None)
            
            # --- INDEXATION DES CLÉS GLOBALES ---
            cles_reclamees = set(df_cmd[df_cmd['Date_Réclamée'].notnull() & (df_cmd['Date_Réclamée'].astype(str).str.strip() != '') & (df_cmd['Date_Réclamée'].astype(str) != 'NaT')]['Clé_unique_détail_commande'].dropna().astype(str))
            cles_facture = set(df_cmd['Clé_unique_détail_facture'].dropna().astype(str)) if 'Clé_unique_détail_facture' in df_cmd.columns else set()
            cles_credite = set(df_cmd[col_cred].dropna().astype(str)) if col_cred else set()
            
            prod_col = '# Produit' if '# Produit' in df_rabais.columns else df_rabais.columns[1]
            debut_col = 'Date début' if 'Date début' in df_rabais.columns else [c for c in df_rabais.columns if 'début' in str(c).lower()][0]
            fin_col = 'Date échéance' if 'Date échéance' in df_rabais.columns else [c for c in df_rabais.columns if 'échéance' in str(c).lower() or 'fin' in str(c).lower()][0]
            rabais_col = 'Rabais' if 'Rabais' in df_rabais.columns else [c for c in df_rabais.columns if 'rabais' in str(c).lower()][0]
            
            df_rabais[debut_col] = pd.to_datetime(df_rabais[debut_col], errors='coerce')
            df_rabais[fin_col] = pd.to_datetime(df_rabais[fin_col], errors='coerce')
            rabais_dict = {prod: group for prod, group in df_rabais.groupby(prod_col)}
            
            # --- PRÉ-CALCUL RAPIDE DE _rabais_calc ---
            rabais_calc_arr = np.zeros(len(df_cmd))
            date_deb_arr = [''] * len(df_cmd)
            date_fin_arr = [''] * len(df_cmd)
            
            for idx, row in df_cmd.iterrows():
                prod = row.get('No_Produit', None)
                df_date = row.get('Date_Facture', pd.NaT)
                qte = row.get('Qté_commandée', 0)
                
                if prod in rabais_dict and pd.notnull(df_date):
                    sub = rabais_dict[prod]
                    valid = sub[(sub[debut_col] <= df_date + pd.Timedelta(days=tolerance)) & (sub[fin_col] >= df_date - pd.Timedelta(days=tolerance))].copy()
                    if not valid.empty:
                        valid['dH'] = (valid[debut_col] - df_date).abs()
                        valid['dI'] = (valid[fin_col] - df_date).abs()
                        valid['dist'] = valid[['dH', 'dI']].min(axis=1)
                        valid = valid.sort_values(by=['dist', debut_col], ascending=[True, False])
                        best = valid.iloc[0]
                        taux = float(best[rabais_col]) if pd.notnull(best[rabais_col]) else 0
                        rabais_calc_arr[idx] = taux * qte
                        date_deb_arr[idx] = best[debut_col].strftime('%Y-%m-%d') if pd.notnull(best[debut_col]) else ''
                        date_fin_arr[idx] = best[fin_col].strftime('%Y-%m-%d') if pd.notnull(best[fin_col]) else ''
            
            df_cmd['_rabais_calc'] = rabais_calc_arr

            # --- PRÉ-CALCUL DE LA COLONNE N ---
            col_n_arr = np.zeros(len(df_cmd))
            for _, group in df_cmd.groupby('Clé_unique_détail_commande'):
                max_r = group['_rabais_calc'].max()
                col_n_arr[group[group['_rabais_calc'] == max_r].index] = 1
            df_cmd['_col_N'] = col_n_arr

            # --- BOUCLE D'ÉCRITURE LIGNE PAR LIGNE SÉCURISÉE ---
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    row = df_cmd.loc[idx]
                    
                    prod = row.get('No_Produit', None)
                    qte = row.get('Qté_commandée', 0)
                    montant_st = row.get('Montant_ST', 0)
                    
                    code_promo = ""
                    if col_promo and col_promo in df_cmd.columns:
                        raw_promo = row.get(col_promo, '')
                        if pd.notnull(raw_promo) and str(raw_promo).strip() != '' and str(raw_promo).lower() != 'nan':
                            code_promo = str(raw_promo).strip()
                    
                    cle_cmd = str(row.get('Clé_unique_détail_commande', ''))
                    if cle_cmd == 'nan' or not cle_cmd:
                        cle_cmd = ""
                        
                    cle_fact = str(row.get('Clé_unique_détail_facture', '')) if pd.notnull(row.get('Clé_unique_détail_facture', '')) else ''
                    if cle_fact == 'nan' or not cle_fact:
                        cle_fact = ""
                        
                    cle_cred_val = str(row.get(col_cred, '')) if col_cred and pd.notnull(row.get(col_cred, '')) else ''
                    date_recl = row.get('Date_Réclamée', None)
                    
                    # --- ÉVALUATION ---
                    suppr_1 = "Supprimer" if pd.notnull(date_recl) and str(date_recl).strip() != "" and str(date_recl) != "NaT" else ""
                    
                    # Colonne C : Valeur de la clé de commande si réclamée
                    val_col_c = cle_cmd if (cle_cmd in cles_reclamees and cle_cmd != "") else ""
                    suppr_d = "Supprimer" if val_col_c != "" else ""
                    
                    cond_suppr_3 = ((cle_cred_val in cles_facture and cle_cred_val != '' and cle_cred_val != 'nan') or 
                                     (cle_fact in cles_credite and cle_fact != '' and cle_fact != 'nan'))
                    suppr_3 = "Supprimer" if cond_suppr_3 else ""
                    
                    # Colonne F : Valeur de la facture si créditée
                    val_col_f = cle_fact if (cle_fact in cles_credite and cle_fact != "") else ""
                    suppr_g = "Supprimer" if val_col_f != "" else ""
                    
                    suppr_5 = "Supprimer" if qte < 0 else ""
                    
                    # Colonne I (Supprimer #6)
                    suppr_6 = ""
                    is_no_recl = (pd.isnull(date_recl) or str(date_recl).strip() == "" or str(date_recl) == "NaT")
                    if is_no_recl:
                        mask_i = (df_cmd['Clé_unique_détail_commande'].astype(str) == cle_cmd) & \
                                 (df_cmd['No_Produit'].astype(str) == str(prod)) & \
                                 (df_cmd['Date_Réclamée'].isnull() | (df_cmd['Date_Réclamée'].astype(str).str.strip() == ""))
                        sub_i = df_cmd[mask_i]
                        if not sub_i.empty and sub_i['Qté_commandée'].sum() >= 0:
                            if row['_rabais_calc'] < sub_i['_rabais_calc'].max():
                                suppr_6 = "Supprimer"
                    
                    # Colonne J (Supprimer #7)
                    suppr_7 = ""
                    val_n = row['_col_N']
                    if val_n == 0:
                        suppr_7 = "Supprimer"
                    else:
                        mask_n1 = (df_cmd['Clé_unique_détail_commande'].astype(str) == cle_cmd) & (df_cmd['_col_N'] == 1)
                        sub_n1 = df_cmd[mask_n1]
                        if not sub_n1.empty:
                            max_facture = sub_n1['Clé_unique_détail_facture'].max()
                            if pd.notnull(row.get('Clé_unique_détail_facture')) and row.get('Clé_unique_détail_facture') < max_facture:
                                suppr_7 = "Supprimer"

                    # Colonne K (Supprimer #8)
                    suppr_8 = ""
                    mask_k = (df_cmd['Clé_unique_détail_commande'].astype(str) == cle_cmd) & (df_cmd['No_Produit'].astype(str) == str(prod))
                    sub_k = df_cmd[mask_k]
                    
                    if not sub_k.empty:
                        cond1 = (sub_k['Qté_commandée'].sum() > 0)
                        cond2 = False
                        if col_date_recl_cred and col_date_recl_cred in df_cmd.columns:
                            d_col = sub_k[col_date_recl_cred]
                            valid_dates = d_col.notnull() & (d_col.astype(str).str.strip() != "") & (d_col.astype(str) != "NaT") & (d_col.astype(str) != "nan")
                            cond2 = valid_dates.any()
                            
                        cond3 = True
                        if col_cred and col_cred in df_cmd.columns:
                            c_col = sub_k[col_cred]
                            valid_creds = c_col.notnull() & (c_col.astype(str).str.strip() != "") & (c_col.astype(str) != "0") & (c_col.astype(str) != "nan")
                            cond3 = (valid_creds.sum() == 0)
                            
                        cond4 = (montant_st < 0.99)
                        
                        if cond1 and cond2 and cond3 and cond4:
                            suppr_8 = "Supprimer"

                    # Colonne L (Supprimer #9)
                    suppr_9 = ""
                    if code_promo.upper().startswith("FIL"):
                        suppr_9 = "Supprimer"

                    suppr_10 = "Supprimer" if montant_st < 0.99 else ""
                    
                    # --- CALCULS FINANCIERS ---
                    rabais_entre_2_dates = row['_rabais_calc']
                    rabais_total = qte * montant_st
                    ecart = rabais_total - rabais_entre_2_dates
                    indicateur_tolerance = 0
                    val_col_n = row['_col_N']
                    
                    # --- SUPPRIMER TOTAL (Col A) ---
                    tous_criteres = [suppr_1, suppr_d, suppr_3, suppr_g, suppr_5, suppr_6, suppr_7, suppr_8, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in tous_criteres) else ""
                    
                    # --- ÉCRITURE DANS LE CLASSEUR EXCEL ---
                    ws_cmd.cell(row=r, column=1).value = suppr_total
                    ws_cmd.cell(row=r, column=2).value = suppr_1
                    ws_cmd.cell(row=r, column=3).value = val_col_c  # Col C : Valeur de la clé
                    ws_cmd.cell(row=r, column=4).value = suppr_d
                    ws_cmd.cell(row=r, column=5).value = suppr_3
                    ws_cmd.cell(row=r, column=6).value = val_col_f  # Col F : Valeur de la facture
                    ws_cmd.cell(row=r, column=7).value = suppr_g
                    ws_cmd.cell(row=r, column=8).value = suppr_5
                    ws_cmd.cell(row=r, column=9).value = suppr_6
                    ws_cmd.cell(row=r, column=10).value = suppr_7
                    ws_cmd.cell(row=r, column=11).value = suppr_8
                    ws_cmd.cell(row=r, column=12).value = suppr_9
                    ws_cmd.cell(row=r, column=13).value = suppr_10
                    
                    ws_cmd.cell(row=r, column=14).value = val_col_n
                    ws_cmd.cell(row=r, column=15).value = rabais_total
                    ws_cmd.cell(row=r, column=16).value = rabais_entre_2_dates
                    ws_cmd.cell(row=r, column=17).value = indicateur_tolerance
                    ws_cmd.cell(row=r, column=18).value = tolerance
                    
                    if date_deb_arr[idx]:
                        ws_cmd.cell(row=r, column=19).value = date_deb_arr[idx]
                    if date_fin_arr[idx]:
                        ws_cmd.cell(row=r, column=20).value = date_fin_arr[idx]
                        
                    ws_cmd.cell(row=r, column=22).value = ecart
            
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            timestamp_str = datetime.now(ZoneInfo("America/Montreal")).strftime("%Y%m%d_%H%M")
            nom_fichier = f"Rapport_Rabais_Final_v{st.session_state.version_compteur}_{timestamp_str}.xlsx"
            
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
