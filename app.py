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
    st.info("Traitement en cours... Injection des formules d'origine.")
    
    try:
        # Chargement du fichier téléversé
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            max_row = ws.max_row
            
            # Dictionnaire des formules extraites directement et fidèlement de votre version de référence d'origine
            formules_ref = {
                1: '=IF(COUNTIF(B2:M2,"Supprimer")=0,"","Supprimer")',
                2: '=IF(INDEX($W2:$DK2,,MATCH("Date_Réclamée",$W$1:$DK$1,0))<>"","Supprimer","")',
                3: '=IF(INDEX($W2:$DK2,,MATCH("Date_Réclamée",$W$1:$DK$1,0))<>"",INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),"")',
                5: '=IF(OR(SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)))<>0,SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)))<>0),"Supprimer","")',
                6: '=IF(SUMIF(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_facture",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)))>0,INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),"")',
                8: '=IF(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Qté_commandée",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)))<0,"Supprimer","")',
                9: '=IF(SUMIFS(INDEX($O$2:$DK$9931,,MATCH("Qté_commandée",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Date_Réclamée",$O$1:$DK$1,0)),"")>=0,IF(INDEX($O2:$DK2,,MATCH("Rabais entre 2 date",$O$1:$DK$1,0))<_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Rabais entre 2 date",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("No_Produit",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("No_Produit",$O$1:$DK$1,0))),"Supprimer",""),"")',
                10: '=IF(N2=0,"Supprimer",IF(INDEX($O2:$DK2,,MATCH("Clé_unique_détail_facture",$O$1:$DK$1,0))<_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_facture",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),$N$2:$N$9931,1),"Supprimer",""))',
                11: '=IF(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Qté_commandée",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)))>0,IF(AND(SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Date_réclamé_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)))<>"",SUMIFS(INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_crédité",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("Clé_unique_détail_commande",$W$1:$DK$1,0)),INDEX($W$2:$DK$9931,,MATCH("No_Produit",$W$1:$DK$1,0)),INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)))=0,INDEX($W2:$DK2,,MATCH("Montant_ST",$W$1:$DK$1,0))<0.99),"Supprimer",""),"")',
                12: '=IF(IFERROR(SEARCH("FIL",INDEX($W2:$DK2,,MATCH("Code_de_Promotion",$W$1:$DK$1,0))),0)=1,"Supprimer","")',
                13: '=IF(INDEX($W2:$DK2,,MATCH("Montant_ST",$W$1:$DK$1,0))<0.99,"Supprimer","")',
                14: '=IF(_xlfn.MAXIFS(INDEX($O$2:$DK$9931,,MATCH("Rabais entre 2 date",$O$1:$DK$1,0)),INDEX($O$2:$DK$9931,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)),INDEX($O2:$DK2,,MATCH("Clé_unique_détail_commande",$O$1:$DK$1,0)))=P2,1,0)',
                15: '=INDEX($W2:$DK2,,MATCH("Qté_commandée",$W$1:$DK$1,0))*INDEX($W2:$DK2,,MATCH("Montant_ST",$W$1:$DK$1,0))',
                16: '=_xlfn.LET(_xlpm.prod,INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)),_xlpm.df,INDEX($W2:$DK2,,MATCH("Date_Facture",$W$1:$DK$1,0)),_xlpm.tol,R2,_xlpm.qte,INDEX($W2:$DK2,,MATCH("Qté_commandée",$W$1:$DK$1,0)),_xlpm.B,\'Rabais entre 2 dates\'!$B$2:$B$20000,_xlpm.H,\'Rabais entre 2 dates\'!$H$2:$H$20000,_xlpm.I,\'Rabais entre 2 dates\'!$I$2:$I$20000,_xlpm.J,\'Rabais entre 2 dates\'!$J$2:$J$20000,_xlpm.crit,(_xlpm.B=_xlpm.prod)*(_xlpm.H<=_xlpm.df+_xlpm.tol)*(_xlpm.I>=_xlpm.df-_xlpm.tol),_xlpm.Hf,_xlfn._xlws.FILTER(_xlpm.H,_xlpm.crit),_xlpm.If,_xlfn._xlws.FILTER(_xlpm.I,_xlpm.crit),_xlpm.Jf,_xlfn._xlws.FILTER(_xlpm.J,_xlpm.crit),_xlpm.dH,ABS(_xlpm.Hf-_xlpm.df),_xlpm.dI,ABS(_xlpm.If-_xlpm.df),_xlpm.dist,(_xlpm.dH+_xlpm.dI-ABS(_xlpm.dH-_xlpm.dI))/2,_xlpm.rab,INDEX(_xlfn.SORTBY(_xlpm.Jf,_xlpm.dist,1,_xlpm.Hf,-1),1),IFERROR(_xlpm.rab*_xlpm.qte,0))',
                17: '=IF(OR(S2="",T2=""),0,IF(AND(INDEX($W2:$DK2,,MATCH("Date_Facture",$W$1:$DK$1,0))>=S2,INDEX($W2:$DK2,,MATCH("Date_Facture",$W$1:$DK$1,0))<=T2),0,1))',
                18: '=R1',
                19: '=IFERROR(_xlfn.LET(_xlpm.prod,INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)),_xlpm.df,INDEX($W2:$DK2,,MATCH("Date_Facture",$W$1:$DK$1,0)),_xlpm.tol,R2,_xlpm.B,\'Rabais entre 2 dates\'!$B$2:$B$20000,_xlpm.H,\'Rabais entre 2 dates\'!$H$2:$H$20000,_xlpm.I,\'Rabais entre 2 dates\'!$I$2:$I$20000,_xlpm.crit,(_xlpm.B=_xlpm.prod)*(_xlpm.H<=_xlpm.df+_xlpm.tol)*(_xlpm.I>=_xlpm.df-_xlpm.tol),_xlpm.Hf,_xlfn._xlws.FILTER(_xlpm.H,_xlpm.crit),_xlpm.If,_xlfn._xlws.FILTER(_xlpm.I,_xlpm.crit),_xlpm.dH,ABS(_xlpm.Hf-_xlpm.df),_xlpm.dI,ABS(_xlpm.If-_xlpm.df),_xlpm.dist,(_xlpm.dH+_xlpm.dI-ABS(_xlpm.dH-_xlpm.dI))/2,INDEX(_xlfn.SORTBY(_xlpm.Hf,_xlpm.dist,1,_xlpm.Hf,-1),1)),"")',
                20: '=IFERROR(_xlfn.LET(_xlpm.prod,INDEX($W2:$DK2,,MATCH("No_Produit",$W$1:$DK$1,0)),_xlpm.df,INDEX($W2:$DK2,,MATCH("Date_Facture",$W$1:$DK$1,0)),_xlpm.tol,R2,_xlpm.B,\'Rabais entre 2 dates\'!$B$2:$B$20000,_xlpm.H,\'Rabais entre 2 dates\'!$H$2:$H$20000,_xlpm.I,\'Rabais entre 2 dates\'!$I$2:$I$20000,_xlpm.crit,(_xlpm.B=_xlpm.prod)*(_xlpm.H<=_xlpm.df+_xlpm.tol)*(_xlpm.I>=_xlpm.df-_xlpm.tol),_xlpm.Hf,_xlfn._xlws.FILTER(_xlpm.H,_xlpm.crit),_xlpm.If,_xlfn._xlws.FILTER(_xlpm.I,_xlpm.crit),_xlpm.dH,ABS(_xlpm.Hf-_xlpm.df),_xlpm.dI,ABS(_xlpm.If-_xlpm.df),_xlpm.dist,(_xlpm.dH+_xlpm.dI-ABS(_xlpm.dH-_xlpm.dI))/2,INDEX(_xlfn.SORTBY(_xlpm.If,_xlpm.dist,1,_xlpm.Hf,-1),1)),"")',
                21: '=IFERROR(_xlfn.XLOOKUP(S2&T2&AB2,\'Rabais entre 2 dates\'!$H$2:$H$20000&\'Rabais entre 2 dates\'!$I$2:$I$20000&\'Rabais entre 2 dates\'!$B$2:$B$20000,\'Rabais entre 2 dates\'!$L$2:$L$20000,""),"")',
                22: '=O2-P2'
            }
            
            def adapt_formula(f_str, r):
                if not f_str.startswith('='):
                    return f_str
                pattern = r'([A-Z]+)(\$?)(2)\b'
                def repl(m):
                    col, dollar, row_num = m.groups()
                    return m.group(0) if dollar == '$' else f"{col}{dollar}{r}"
                return re.sub(pattern, repl, f_str)

            # Application rigoureuse de la ligne 2 jusqu'à la dernière ligne dynamique
            for r in range(2, max_row + 1):
                for c in range(1, 23):
                    if c in formules_ref:
                        ws.cell(row=r, column=c).value = adapt_formula(formules_ref[c], r)
                    else:
                        ws.cell(row=r, column=c).value = None

            # Sauvegarde en mémoire du fichier final
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées avec exactitude.")
            
            # 3. BOUTON DE TÉLÉCHARGEMENT
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
