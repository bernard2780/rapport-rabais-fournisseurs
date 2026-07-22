import streamlit as st
import pandas as pd
import openpyxl
import io
import re

# 1. CONFIGURATION DE LA PAGE WEB
st.set_page_config(page_title="Générateur de Rapport", layout="wide")
st.title("Générateur de Rapport : Rabais Fournisseurs")
st.write("Veuillez téléverser votre fichier d'inventaire brut ci-dessous.")

# 2. BOUTON D'IMPORTATION
fichier_upload = st.file_uploader("Choisissez le fichier de commandes (.xlsx)", type=["xlsx"])

if fichier_upload is not None:
    st.info("Traitement en cours... Extraction et application dynamique des formules d'origine.")
    
    try:
        # Chargement du fichier téléversé
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            max_row = ws.max_row
            
            # Extraction des formules exactes depuis la ligne 2 du fichier
            formules_ligne_2 = [ws.cell(row=2, column=col).value for col in range(1, 23)]
            
            # Fonction d'adaptation dynamique (remplace la ligne 2 par la ligne r sans toucher aux références absolues $)
            def adapt_formula(formula, r):
                if formula is None:
                    return None
                f_str = formula.text if hasattr(formula, 'text') else str(formula)
                if not f_str.startswith('='):
                    return f_str
                
                # Expression régulière pour remplacer les numéros de ligne relatifs (ex: B2 -> Br, mais pas $B$2)
                pattern = r'([A-Z]+)(\$?)(2)\b'
                def repl(m):
                    col, dollar, row_num = m.groups()
                    return m.group(0) if dollar == '$' else f"{col}{dollar}{r}"
                
                return re.sub(pattern, repl, f_str)

            # Application rigoureuse de la ligne 2 jusqu'à la dernière ligne dynamique
            for r in range(2, max_row + 1):
                for idx, formule in enumerate(formules_ligne_2):
                    ws.cell(row=r, column=idx + 1).value = adapt_formula(formule, r)

            # Sauvegarde en mémoire du fichier final
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées avec exactitude.")
            
            # 4. BOUTON DE TÉLÉCHARGEMENT
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
