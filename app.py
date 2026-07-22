import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs")
st.write("Veuillez téléverser votre fichier brut d'inventaire ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Fichier reçu. Traitement et application des règles de calcul en cours...")
    
    try:
        # 3. LECTURE DES DONNÉES
        df_commandes = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
        df_dates = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
        
        if not df_commandes.empty:
            # Ligne de paramètres (Ligne 2 d'Excel, index 0) et données (à partir de la ligne 3)
            ligne_2_parametres = df_commandes.iloc[0]
            df_donnees = df_commandes.iloc[1:].copy().reset_index(drop=True)
            
            # 4. MOTEUR DE CALCUL FIDÈLE AUX RÈGLES EXCEL
            
            # Nettoyage et conversion numérique
            if 'Montant_ST' in df_donnees.columns:
                df_donnees['Montant_ST'] = pd.to_numeric(df_donnees['Montant_ST'], errors='coerce')
            if 'Qté_commandée' in df_donnees.columns:
                df_donnees['Qté_commandée'] = pd.to_numeric(df_donnees['Qté_commandée'], errors='coerce')

            # Règle #9 : Code de promotion contenant "FIL"
            if 'Code_de_Promotion' in df_donnees.columns:
                df_donnees['Supprimer #9'] = np.where(
                    df_donnees['Code_de_Promotion'].astype(str).str.contains('FIL', na=False), 
                    'Supprimer', ''
                )
            
            # Règle #10 : Montant inférieur à 0.99$
            if 'Montant_ST' in df_donnees.columns:
                df_donnees['Supprimer #10'] = np.where(
                    df_donnees['Montant_ST'] < 0.99, 
                    'Supprimer', ''
                )
            
            # Calcul du Rabais total (Qté x Montant_ST)
            if 'Qté_commandée' in df_donnees.columns and 'Montant_ST' in df_donnees.columns:
                df_donnees['Rabais total'] = df_donnees['Qté_commandée'] * df_donnees['Montant_ST']
            
            # Consolidation du Supprimer total (si l'une des colonnes de suppression contient "Supprimer")
            colonnes_supprimer = [col for col in df_donnees.columns if 'Supprimer #' in str(col)]
            if colonnes_supprimer:
                df_donnees['Supprimer total'] = np.where(
                    df_donnees[colonnes_supprimer].eq('Supprimer').any(axis=1), 
                    'Supprimer', ''
                )
            else:
                df_donnees['Supprimer total'] = ''
                
        st.success("Traitement terminé avec succès ! Les résultats sont alignés sur vos critères.")
        
        # 5. AFFICHAGE DES RÉSULTATS
        st.subheader("Aperçu des résultats (5 premières lignes)")
        st.dataframe(df_donnees.head())
        
        # 6. GÉNÉRATION DU NOUVEAU FICHIER EXCEL TÉLÉCHARGEABLE
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_donnees.to_excel(writer, sheet_name='Rapport Final', index=False)
        
        st.download_button(
            label="📥 Télécharger le rapport final calculé (Excel)",
            data=buffer.getvalue(),
            file_name="Rapport_Rabais_Calcule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
