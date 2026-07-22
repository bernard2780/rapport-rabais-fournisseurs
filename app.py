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
    st.info("Fichier reçu. Injection des formules et traitement dynamique en cours...")
    
    try:
        # 3. CHARGEMENT DU FICHIER AVEC OPENPYXL
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            
            # Détection dynamique de la dernière ligne de données
            max_row = ws.max_row
            
            # Injection dynamique des formules ligne par ligne (de la ligne 2 jusqu'à la fin)
            for r in range(2, max_row + 1):
                # Colonne A (Supprimer total)
                ws.cell(row=r, column=1).value = f'=IF(COUNTIF(B{r}:M{r},"Supprimer")=0,"","Supprimer")'
                
                # Colonnes B à M (Règles de suppression #1 à #10)
                ws.cell(row=r, column=2).value = f'=IF(INDEX($W{r}:$DK{r},,MATCH("Date_Réclamée",$W$1:$DK$1,0))<>"","Supprimer","")'
                ws.cell(row=r, column=3).value = f'=IF(INDEX($W{r}:$DK{r},,MATCH("Date_Réclamée",$W$1:$DK$1,0))<>"",INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),"")'
                ws.cell(row=r, column=5).value = f'=IF(OR(SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)))<>0,SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)))<>0),"Supprimer","")'
                ws.cell(row=r, column=6).value = f'=IF(SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)))>0,INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),"")'
                ws.cell(row=r, column=8).value = f'=IF(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Qté_commandée",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("No_Produit",$W$1:$DK$1,0)))<0,"Supprimer","")'
                ws.cell(row=r, column=9).value = f'=IF(SUMIFS(INDEX($O$2:$DK$9931,,MATCH("Qté_commandée",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Date_Réclamée",$O$1:$DK$1,0)),"")>=0,IF(INDEX($O{r}:$DK{r},,MATCH("Rabais entre 2 date",$O$1:$DK$1,0))<_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Rabais entre 2 date",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("No_Produit",$O$1:$DK$1,0))),"Supprimer",""),"")'
                ws.cell(row=r, column=10).value = f'=IF(N{r}=0,"Supprimer",IF(INDEX($O{r}:$DK{r},,MATCH("Clé_unique_détail_facture",$O$1:$DK$1,0))<_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),$N$2:$N$9931,1),"Supprimer",""))'
                ws.cell(row=r, column=11).value = f'=IF(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Qté_commandée",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("No_Produit",$W$1:$DK$1,0)))>0,IF(AND(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Date_réclamé_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("No_Produit",$W$1:$DK$1,0)))<>"",SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W{r}:$DK{r},,MATCH("No_Produit",$W$1:$DK$1,0)))=0,INDEX($W{r}:$DK{r},,MATCH("Montant_ST",$W$1:$DK$1,0))<0.99),"Supprimer",""),"")'
                ws.cell(row=r, column=12).value = f'=IF(IFERROR(SEARCH("FIL",INDEX($W{r}:$DK{r},,MATCH("Code_de_Promotion",$W$1:$DK$1,0))),0)=1,"Supprimer","")'
                ws.cell(row=r, column=13).value = f'=IF(INDEX($W{r}:$DK{r},,MATCH("Montant_ST",$W$1:$DK$1,0))<0.99,"Supprimer","")'
                
                # Colonnes N à V (Calculs des rabais)
                ws.cell(row=r, column=14).value = f'=IF(_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Rabais entre 2 date",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O{r}:$DK{r},,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)))=P{r},1,0)'
                ws.cell(row=r, column=15).value = f'=INDEX($W{r}:$DK{r},,,MATCH("Qté_commandée",$W$1:$DK$1,0))*INDEX($W{r}:$DK{r},,MATCH("Montant_ST",$W$1:$DK$1,0))'
                ws.cell(row=r, column=17).value = f'=IF(OR(S{r}="",T{r}=""),0,IF(AND(INDEX($W{r}:$DK{r},,MATCH("Date_Facture",$W$1:$DK$1,0))>=S{r},INDEX($W{r}:$DK{r},,MATCH("Date_Facture",$W$1:$DK$1,0))<=T{r}),0,1))'
                ws.cell(row=r, column=18).value = f'=R1'
                ws.cell(row=r, column=22).value = f'=O{r}-P{r}'

            # Sauvegarde en mémoire
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes calculées et injectées dynamiquement.")
            
            # Bouton de téléchargement
            st.download_button(
                label="📥 Télécharger le rapport final complet (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("L'onglet 'Rabais fournisseurs' est introuvable.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
