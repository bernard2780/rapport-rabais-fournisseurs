import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. CONFIGURATION DE LA PAGE
st.set_page_config(page_title="Générateur de Rapport", layout="wide")

st.title("Générateur de Rapport : Rabais Fournisseurs")

# 2. SÉCURITÉ : MOT DE PASSE
mot_de_passe = st.text_input("Veuillez entrer le mot de passe de l'entreprise pour accéder à l'outil :", type="password")

if mot_de_passe != "Fillion2026!":
    if mot_de_passe: # Si quelqu'un tape un mauvais mot de passe
        st.error("Mot de passe incorrect. L'accès est refusé.")
    st.stop() # Cette commande agit comme un mur : elle empêche l'exécution de la suite du code

# 3. L'APPLICATION WEB (Accessible uniquement si le mot de passe est bon)
st.write("Veuillez téléverser votre fichier brut d'inventaire (sans formules) ci-dessous.")

# Création du bouton d'importation
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Fichier reçu. Croisement des données en cours, veuillez patienter...")
    
    try:
        # LECTURE DES DONNÉES EN MÉMOIRE
        df_commandes = pd.read_excel(fichier_upload, sheet_name='Rabais fournisseurs')
        df_dates = pd.read_excel(fichier_upload, sheet_name='Rabais entre 2 dates')
        
        # MOTEUR DE CALCUL MATHÉMATIQUE
        # Règle M (Supprimer #10)
        if 'Montant_ST' in df_commandes.columns:
            df_commandes['Supprimer #10'] = np.where(df_commandes['Montant_ST'] < 0.99, 'Supprimer', '')
        
        # Règle L (Supprimer #9)
        if 'Code_de_Promotion' in df_commandes.columns:
            df_commandes['Supprimer #9'] = np.where(df_commandes['Code_de_Promotion'].astype(str).str.contains('FIL', na=False), 'Supprimer', '')
            
        # Règle O (Rabais total)
        if 'Qté_commandée' in df_commandes.columns and 'Montant_ST' in df_commandes.columns:
            df_commandes['Rabais total'] = df_commandes['Qté_commandée'] * df_commandes['Montant_ST']
            
        # Règle A (Supprimer total)
        colonnes_supprimer = [col for col in df_commandes.columns if 'Supprimer #' in str(col)]
        df_commandes['Supprimer total'] = np.where(df_commandes[colonnes_supprimer].eq('Supprimer').any(axis=1), 'Supprimer', '')
        
        st.success("Calculs terminés avec succès !")
        
        # AFFICHAGE DES RÉSULTATS
        st.subheader("Aperçu des résultats (5 premières lignes)")
        st.dataframe(df_commandes.head())
        
        # GÉNÉRATION DU FICHIER TÉLÉCHARGEABLE
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            df_commandes.to_excel(writer, sheet_name='Rapport Final', index=False)
        
        st.download_button(
            label="📥 Télécharger le rapport final calculé (Excel)",
            data=buffer.getvalue(),
            file_name="Rapport_Rabais_Calcule.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    except Exception as e:
        st.error(f"Une erreur s'est produite lors de la lecture des colonnes : {e}")
