import streamlit as st
import pandas as pd
import openpyxl
import io

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs")
st.write("Veuillez téléverser votre fichier d'inventaire ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Fichier reçu. Injection dynamique des formules et traitement en cours...")
    
    try:
        # 3. CHARGEMENT DU FICHIER AVEC OPENPYXL POUR PRÉSERVER LES FORMULES
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            
            # Détection dynamique de la dernière ligne de données, sans aucune limite fixe
            max_row = ws.max_row
            
            # Propagation des formules de la ligne 2 jusqu'à la dernière ligne dynamique pour les colonnes A à V (1 à 22)
            for col in range(1, 23):
                formule_ligne_2 = ws.cell(row=2, column=col).value
                if formule_ligne_2:
                    for r in range(3, max_row + 1):
                        ws.cell(row=r, column=col).value = formule_ligne_2
            
            # Sauvegarde en mémoire du fichier mis à jour avec toutes les formules propagées
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées dynamiquement.")
            
            # 4. BOUTON DE TÉLÉCHARGEMENT DU RAPPORT FINAL
            st.download_button(
                label="📥 Télécharger le rapport final avec formules (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("L'onglet 'Rabais fournisseurs' est introuvable dans le fichier téléversé.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite lors du traitement du fichier : {e}")
