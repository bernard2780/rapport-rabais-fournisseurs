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
    st.info(f"Traitement rigoureux et complet (Version {st.session_state.version_compteur}) en cours...")
    
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
            
            # --- DÉTECTION SÉCURISÉE DES COLONNES PAR MOTS-CLÉS ---
            def trouve_col(mots_cles, default=None):
                for col in df_cmd.columns:
                    col_lower = str(col).lower()
                    if all(m in col_lower for m in mots_cles):
                        return col
                return default

            col_cle_cmd = trouve_col(['clé', 'commande'], df_cmd.columns[0])
            col_produit = trouve_col(['produit'], df_cmd.columns[1])
            col_qte = trouve_col(['qte', 'quantité'], 'Qté_commandée' if 'Qté_commandée' in df_cmd.columns else df_cmd.columns[2])
            col_montant = trouve_col(['montant', 'st'], 'Montant_ST' if 'Montant_ST' in df_cmd.columns else df_cmd.columns[3])
            col_date_fact = trouve_col(['date', 'facture'], 'Date_Facture' if 'Date_Facture' in df_cmd.columns else None)
            col_date_recl = trouve_col(['réclamée'], 'Date_Réclamée' if 'Date_Réclamée' in df_cmd.columns else None)
            col_cle_fact = trouve_col(['clé', 'facture'], 'Clé_unique_détail_facture' if 'Clé_unique_détail_facture' in df_cmd.columns else None)
            col_cred = trouve_col(['crédit', 'clé'], None)
            col_date_recl_cred = trouve_col(['crédit', 'date'], None)
            
            col_promo = None
            for candidate in ['Code_de_Promotion', 'Code_promotion', 'Code promotion', 'Code de Promotion']:
                if candidate in df_cmd.columns:
                    col_promo = candidate
                    break
            if not col_promo:
                col_promo = next((c for c in df_cmd.columns if 'promo' in str(c).lower()), None)

            # Nettoyage des types numériques et dates
            df_cmd['__qte_num'] = pd.to_numeric(df_cmd[col_qte], errors='coerce').fillna(0)
            df_cmd['__montant_num'] = pd.to_numeric(df_cmd[col_montant], errors='coerce').fillna(0)
            if col_date_fact and col_date_fact in df_cmd.columns:
                df_cmd['__date_f'] = pd.to_datetime(df_cmd[col_date_fact], errors='coerce')
            else:
                df_cmd['__date_f'] = pd.NaT

            # --- 1. CALCUL DES RABAIS ENTRE 2 DATES ---
            prod_col_rab = '# Produit' if '# Produit' in df_rabais.columns else df_rabais.columns[1]
            debut_col = 'Date début' if 'Date début' in df_rabais.columns else [c for c in df_rabais.columns if 'début' in str(c).lower()][0]
            fin_col = 'Date échéance' if 'Date échéance' in df_rabais.columns else [c for c in df_rabais.columns if 'échéance' in str(c).lower() or 'fin' in str(c).lower()][0]
            rabais_col = 'Rabais' if 'Rabais' in df_rabais.columns else [c for c in df_rabais.columns if 'rabais' in str(c).lower()][0]
            
            df_rabais[debut_col] = pd.to_datetime(df_rabais[debut_col], errors='coerce')
            df_rabais[fin_col] = pd.to_datetime(df_rabais[fin_col], errors='coerce')
            rabais_dict = {prod: group for prod, group in df_rabais.groupby(prod_col_rab)}
            
            rabais_calc_arr = np.zeros(len(df_cmd))
            date_deb_arr = [''] * len(df_cmd)
            date_fin_arr = [''] * len(df_cmd)
            
            for idx, row in df_cmd.iterrows():
                prod = row.get(col_produit, None)
                df_date = row.get('__date_f', pd.NaT)
                qte = row.get('__qte_num', 0)
                
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

            # --- 2. COLONNE N ---
            col_n_arr = np.zeros(len(df_cmd))
            for _, group in df_cmd.groupby(col_cle_cmd):
                max_r = group['_rabais_calc'].max()
                col_n_arr[group[group['_rabais_calc'] == max_r].index] = 1
            df_cmd['_col_N'] = col_n_arr

            # --- 3. INDEXATION RECHERCHEX ---
            h_col = df_rabais.columns[7] if len(df_rabais.columns) > 7 else debut_col
            i_col = df_rabais.columns[8] if len(df_rabais.columns) > 8 else fin_col
            b_col = df_rabais.columns[1] if len(df_rabais.columns) > 1 else prod_col_rab
            l_col = df_rabais.columns[11] if len(df_rabais.columns) > 11 else rabais_col

            rabais_lookup = {}
            for _, r_row in df_rabais.iterrows():
                val_h = pd.to_datetime(r_row.get(h_col), errors='coerce')
                val_i = pd.to_datetime(r_row.get(i_col), errors='coerce')
                val_b = str(r_row.get(b_col, '')).strip()
                val_l = r_row.get(l_col, '')
                if pd.notnull(val_h) and pd.notnull(val_i):
                    rabais_lookup[f"{val_h.strftime('%Y-%m-%d')}{val_i.strftime('%Y-%m-%d')}{val_b}"] = val_l

            # --- 4. ENSEMBLES ET PRÉ-CALCULS POUR COLONNES I et J ---
            series_date_recl = df_cmd[col_date_recl] if col_date_recl and col_date_recl in df_cmd.columns else pd.Series(np.nan, index=df_cmd.index)
            cles_reclamees = set(df_cmd[series_date_recl.notnull() & (series_date_recl.astype(str).str.strip() != '') & (series_date_recl.astype(str) != 'NaT')][col_cle_cmd].dropna().astype(str))
            
            series_cle_fact = df_cmd[col_cle_fact] if col_cle_fact and col_cle_fact in df_cmd.columns else pd.Series(np.nan, index=df_cmd.index)
            cles_facture = set(series_cle_fact.dropna().astype(str))
            
            series_cred = df_cmd[col_cred] if col_cred and col_cred in df_cmd.columns else pd.Series(np.nan, index=df_cmd.index)
            cles_credite = set(series_cred.dropna().astype(str))

            # Pré-calculs pour Col I (Supprimer #6)
            is_no_recl_mask = series_date_recl.isnull() | (series_date_recl.astype(str).str.strip() == "") | (series_date_recl.astype(str) == "NaT")
            df_cmd['_temp_i_rab'] = np.where(is_no_recl_mask, df_cmd['_rabais_calc'], -np.inf)
            group_i_keys = [col_cle_cmd, col_produit]
            max_rab_i_group = df_cmd.groupby(group_i_keys)['_temp_i_rab'].transform('max')
            sum_qte_i_group = df_cmd.groupby(group_i_keys)['__qte_num'].transform('sum')

            # Pré-calculs pour Col J (Supprimer #7)
            max_fact_n1_filled = pd.Series(np.nan, index=df_cmd.index)
            if col_cle_fact and col_cle_fact in df_cmd.columns:
                n1_mask = df_cmd['_col_N'] == 1
                df_cmd['_max_fact_n1'] = np.where(n1_mask, df_cmd[col_cle_fact], np.nan)
                max_fact_n1_filled = df_cmd.groupby(col_cle_cmd)['_max_fact_n1'].transform('max')

            # --- 5. ÉCRITURE DANS OPENPYXL ---
            for r in range(2, max_row + 1):
                idx = r - 2
                if idx < len(df_cmd):
                    row = df_cmd.loc[idx]
                    qte = row['__qte_num']
                    montant_st = row['__montant_num']
                    
                    d_s = date_deb_arr[idx]
                    d_t = date_fin_arr[idx]
                    val_ab = str(row.get(col_produit, '')).strip()
                    
                    code_promo_val = rabais_lookup.get(f"{d_s}{d_t}{val_ab}", "") if (d_s and d_t) else ""
                    code_promo_str = str(code_promo_val).strip() if pd.notnull(code_promo_val) else ""

                    cle_cmd = str(row.get(col_cle_cmd, ''))
                    if cle_cmd == 'nan' or not cle_cmd: cle_cmd = ""
                    
                    cle_fact = str(row.get(col_cle_fact, '')) if col_cle_fact and pd.notnull(row.get(col_cle_fact, '')) else ''
                    if cle_fact == 'nan' or not cle_fact: cle_fact = ""
                    
                    cle_cred_val = str(row.get(col_cred, '')) if col_cred and pd.notnull(row.get(col_cred, '')) else ''
                    date_recl = row.get(col_date_recl, None) if col_date_recl else None
                    
                    # Règle 1 (Col B)
                    suppr_1 = "Supprimer" if pd.notnull(date_recl) and str(date_recl).strip() != "" and str(date_recl) != "NaT" else ""
                    
                    # Règle 2 / Col C
                    val_col_c = cle_cmd if (cle_cmd in cles_reclamees and cle_cmd != "") else ""
                    suppr_d = "Supprimer" if val_col_c != "" else ""
                    
                    # Règle 3 / Col E
                    cond_suppr_3 = ((cle_cred_val in cles_facture and cle_cred_val != '' and cle_cred_val != 'nan') or 
                                     (cle_fact in cles_credite and cle_fact != '' and cle_fact != 'nan'))
                    suppr_3 = "Supprimer" if cond_suppr_3 else ""
                    
                    # Règle 4 / Col F & G
                    cond_credit_remplie = (cle_fact in cles_credite and cle_fact != '' and cle_fact != 'nan')
                    val_col_f = cle_cmd if cond_credit_remplie else ""
                    suppr_g = "Supprimer" if val_col_f != "" else ""
                    
                    # Règle 5 / Col H
                    suppr_5 = "Supprimer" if qte < 0 else ""
                    
                    # --- COLONNE I (Supprimer #6) ---
                    suppr_6 = ""
                    is_no_recl = (pd.isnull(date_recl) or str(date_recl).strip() == "" or str(date_recl) == "NaT")
                    if is_no_recl:
                        if sum_qte_i_group.iloc[idx] >= 0:
                            if row['_rabais_calc'] < max_rab_i_group.iloc[idx]:
                                suppr_6 = "Supprimer"
                    
                    # --- COLONNE J (Supprimer #7) ---
                    suppr_7 = ""
                    val_n = row['_col_N']
                    if val_n == 0:
                        suppr_7 = "Supprimer"
                    else:
                        m_f = max_fact_n1_filled.iloc[idx]
                        if col_cle_fact and pd.notnull(row.get(col_cle_fact)) and pd.notnull(m_f) and row.get(col_cle_fact) < m_f:
                            suppr_7 = "Supprimer"

                    # Règle 8 / Col K
                    suppr_8 = ""

                    # Règle 9 / Col L
                    suppr_9 = "Supprimer" if code_promo_str.upper().startswith("FIL") else ""

                    # Règle 10 / Col M
                    suppr_10 = "Supprimer" if montant_st < 0.99 else ""
                    
                    rabais_entre_2_dates = row['_rabais_calc']
                    rabais_total = qte * montant_st
                    ecart = rabais_total - rabais_entre_2_dates
                    val_col_n = row['_col_N']
                    
                    tous_criteres = [suppr_1, suppr_d, suppr_3, suppr_g, suppr_5, suppr_6, suppr_7, suppr_8, suppr_9, suppr_10]
                    suppr_total = "Supprimer" if any(c == "Supprimer" for c in tous_criteres) else ""
                    
                    ws_cmd.cell(row=r, column=1).value = suppr_total
                    ws_cmd.cell(row=r, column=2).value = suppr_1
                    ws_cmd.cell(row=r, column=3).value = val_col_c
                    ws_cmd.cell(row=r, column=4).value = suppr_d
                    ws_cmd.cell(row=r, column=5).value = suppr_3
                    ws_cmd.cell(row=r, column=6).value = val_col_f
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
                    ws_cmd.cell(row=r, column=17).value = 0
                    ws_cmd.cell(row=r, column=18).value = tolerance
                    
                    if date_deb_arr[idx]: ws_cmd.cell(row=r, column=19).value = date_deb_arr[idx]
                    if date_fin_arr[idx]: ws_cmd.cell(row=r, column=20).value = date_fin_arr[idx]
                        
                    ws_cmd.cell(row=r, column=21).value = code_promo_str
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
