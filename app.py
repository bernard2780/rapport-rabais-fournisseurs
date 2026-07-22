import streamlit as st
import pandas as pd
import openpyxl
import io

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs")
st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Fichier reçu. Injection et propagation dynamique des formules en cours...")
    
    try:
        # 3. CHARGEMENT DU FICHIER TÉLÉVERSÉ
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            
            # Détection dynamique de la dernière ligne de données
            max_row = ws.max_row
            
            # Récupération de toutes les formules exactes de la ligne 2 (colonnes A à V, soit 1 à 22)
            formules_ligne_2 = [ws.cell(row=2, column=col).value for col in range(1, 23)]
            
            # Propagation rigoureuse de la ligne 3 jusqu'à la dernière ligne dynamique
            for r in range(3, max_row + 1):
                for idx, formule in enumerate(formules_ligne_2):
                    if formule is not None:
                        formule_str = str(formule)
                        
                        # Remplacement dynamique des références de la ligne 2 par la ligne courante (r)
                        for c_lettre in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V']:
                            formule_str = formule_str.replace(f'{c_lettre}2', f'{c_lettre}{r}')
                        
                        ws.cell(row=r, column=idx + 1).value = formule_str
                    else:
                        ws.cell(row=r, column=idx + 1).value = None

            # Sauvegarde en mémoire du fichier mis à jour avec toutes les formules propagées
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées avec l'ensemble des formules.")
            
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
        st.error(f"Une erreur s'est produite lors du traitement : {e}")
