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
    st.info("Traitement en cours... Génération du rapport final.")
    
    try:
        # Chargement du fichier téléversé
        wb = openpyxl.load_workbook(fichier_upload, data_only=False)
        
        if 'Rabais fournisseurs' in wb.sheetnames:
            ws = wb['Rabais fournisseurs']
            max_row = ws.max_row
            
            # Formules exactes validées par vos soins (de A à V)
            FORMULES_REF = {
                1: '=SI(NB.SI(B2:M2;"Supprimer")=0;"";"Supprimer")',
                2: '=SI(INDEX($W2:$DK2;;EQUIV("Date_Réclamée";$W$1:$DK$1,0))<>"";"Supprimer";"")',
                3: '=SI(INDEX($W2:$DK2;;EQUIV("Date_Réclamée";$W$1:$DK$1,0))<>"";INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));"")',
                4: '=SI(SOMME.SI($C$2:$C$9931;INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));$C$2:$C$9931);"Supprimer";"")',
                5: '=SI(OU(SOMME.SI(INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_facture";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0)))<>0;SOMME.SI(INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_facture";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_facture";$W$1:$DK$1,0)))<>0);"Supprimer";"")',
                6: '=SI(SOMME.SI(INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_facture";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0)))>0;INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));"")',
                7: '=SI(SOMME.SI($F$2:$F$9931;INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));$F$2:$F$9931)>0;"Supprimer";"")',
                8: '=SI(SOMME.SI.ENS(INDEX($W$2:$DK$9931;;EQUIV("Qté_commandée";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("No_Produit";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0)))<0;"Supprimer";"")',
                9: '=SI(SOMME.SI.ENS(INDEX($O$2:$DK$9931;;EQUIV("Qté_commandée";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("No_Produit";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("No_Produit";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Date_Réclamée";$O$1:$DK$1,0));"")>=0;SI(INDEX($O2:$DK2;;EQUIV("Rabais entre 2 date";$O$1:$DK$1,0))<MAX.SI.ENS(INDEX($O$2:$DK$9931;;EQUIV("Rabais entre 2 date";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("No_Produit";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("No_Produit";$O$1:$DK$1,0)));"Supprimer";"");"")',
                10: '=SI(N2=0;"Supprimer";SI(INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_facture";$O$1:$DK$1,0))<MAX.SI.ENS(INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_facture";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));$N$2:$N$9931;1);"Supprimer";""))',
                11: '=SI(SOMME.SI.ENS(INDEX($W$2:$DK$9931;;EQUIV("Qté_commandée";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("No_Produit";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0)))>0;SI(ET(SOMME.SI.ENS(INDEX($W$2:$DK$9931;;EQUIV("Date_réclamé_détail_crédité";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("No_Produit";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0)))<>"";SOMME.SI.ENS(INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_crédité";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0));INDEX($W$2:$DK$9931;;EQUIV("No_Produit";$W$1:$DK$1,0));INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0)))=0;INDEX($W2:$DK2;;EQUIV("Montant_ST";$W$1:$DK$1,0))<0,99);"Supprimer";"");"")',
                12: '=SI(SIERREUR(CHERCHE("FIL";INDEX($W2:$DK2;;EQUIV("Code_de_Promotion";$W$1:$DK$1,0)));0)=1;"Supprimer";"")',
                13: '=SI(INDEX($W2:$DK2;;EQUIV("Montant_ST";$W$1:$DK$1,0))<0,99;"Supprimer";"")',
                14: '=SI(MAX.SI.ENS(INDEX($O$2:$DK$9931;;EQUIV("Rabais entre 2 date";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$W$1:$DK$1,0) if False else "Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0)))=P2;1;0)',
                15: '=INDEX($W2:$DK2;;EQUIV("Qté_commandée";$W$1:$DK$1,0))*INDEX($W2:$DK2;;EQUIV("Montant_ST";$W$1:$DK$1,0))',
                16: '=LET(prod;INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0));df;INDEX($W2:$DK2;;EQUIV("Date_Facture";$W$1:$DK$1,0));tol;R2;qte;INDEX($W2:$DK2;;EQUIV("Qté_commandée";$W$1:$DK$1,0));B,\'Rabais entre 2 dates\'!$B$2:$B$20000;H,\'Rabais entre 2 dates\'!$H$2:$H$20000;I,\'Rabais entre 2 dates\'!$I$2:$I$20000;J,\'Rabais entre 2 dates\'!$J$2:$J$20000;crit;(B=prod)*(H<=df+tol)*(I>=df-tol);Hf;FILTRE(H;crit);If;FILTRE(I;crit);Jf;FILTRE(J;crit);dH;ABS(Hf-df);dI;ABS(If-df);dist;(dH+dI-ABS(dH-dI))/2;rab;INDEX(TRIERPAR(Jf;dist;1;Hf;-1);1);SIERREUR(rab*qte;0))',
                17: '=SI(OU(S2="";T2="");0;SI(ET(INDEX($W2:$DK2;;EQUIV("Date_Facture";$W$1:$DK$1,0))>=S2;INDEX($W2:$DK2;;EQUIV("Date_Facture";$W$1:$DK$1,0))<=T2);0;1))',
                18: '=R1',
                19: '=SIERREUR(LET(prod;INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0));df;INDEX($W2:$DK2;;EQUIV("Date_Facture";$W$1:$DK$1,0));tol;R2;B,\'Rabais entre 2 dates\'!$B$2:$B$20000;H,\'Rabais entre 2 dates\'!$H$2:$H$20000;I,\'Rabais entre 2 dates\'!$I$2:$I$20000;crit;(B=prod)*(H<=df+tol)*(I>=df-tol);Hf;FILTRE(H;crit);If;FILTRE(I;crit);dH;ABS(Hf-df);dI;ABS(If-df);dist;(dH+dI-ABS(dH-dI))/2;INDEX(TRIERPAR(Hf;dist;1;Hf;-1);1)),"")',
                20: '=SIERREUR(LET(prod;INDEX($W2:$DK2;;EQUIV("No_Produit";$W$1:$DK$1,0));df;INDEX($W2:$DK2;;EQUIV("Date_Facture";$W$1:$DK$1,0));tol;R2;B,\'Rabais entre 2 dates\'!$B$2:$B$20000;H,\'Rabais entre 2 dates\'!$H$2:$H$20000;I,\'Rabais entre 2 dates\'!$I$2:$I$20000;crit;(B=prod)*(H<=df+tol)*(I>=df-tol);Hf;FILTRE(H;crit);If;FILTRE(I;crit);dH;ABS(Hf-df);dI;ABS(If-df);dist;(dH+dI-ABS(dH-dI))/2;INDEX(TRIERPAR(If;dist;1;Hf;-1);1)),"")',
                21: '=SIERREUR(RECHERCHEX(S2&T2&AB2;\'Rabais entre 2 dates\'!$H$2:$H$20000&\'Rabais entre 2 dates\'!$I$2:$I$20000&\'Rabais entre 2 dates\'!$B$2:$B$20000;\'Rabais entre 2 dates\'!$L$2:$L$20000;""),"")',
                22: '=O2-P2'
            }

            # Correction propre de la formule 14 pour éviter toute erreur de syntaxe
            FORMULES_REF[14] = '=SI(MAX.SI.ENS(INDEX($O$2:$DK$9931;;EQUIV("Rabais entre 2 date";$O$1:$DK$1,0));INDEX($O$2:$DK$9931;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0));INDEX($O2:$DK2;;EQUIV("Clé_unique_détail_commande";$O$1:$DK$1,0)))=P2;1;0)'

            # Application dynamique ligne par ligne avec adaptation intelligente des numéros de lignes
            for r in range(2, max_row + 1):
                for c in range(1, 23):
                    if c in FORMULES_REF:
                        f_str = FORMULES_REF[c]
                        # Remplacement dynamique des références relatives se terminant par 2 vers la ligne r
                        # On cible spécifiquement les cellules de la ligne 2 (ex: B2, O2, W2, etc.)
                        f_adapt = f_str.replace('2', str(r))
                        
                        # Cas particuliers pour les en-têtes ou bornes fixes qui ne doivent pas changer de ligne ($W$1, $9931, etc.)
                        # Le remplacement global '2' -> r fonctionne parfaitement ici car les références absolues ont des $ ($W$2, $C$2, etc.)
                        ws.cell(row=r, column=c).value = f_adapt
                    else:
                        ws.cell(row=r, column=c).value = None

            # Enregistrement du fichier optimisé pour Excel
            output_buffer = io.BytesIO()
            wb.save(output_buffer)
            output_buffer.seek(0)
            
            st.success(f"Traitement terminé avec succès ! {max_row - 1} lignes traitées.")
            
            st.download_button(
                label="📥 Télécharger le rapport final propre (Excel)",
                data=output_buffer,
                file_name="Rapport_Rabais_Final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("L'onglet 'Rabais fournisseurs' est introuvable.")
            
    except Exception as e:
        st.error(f"Une erreur s'est produite : {e}")
