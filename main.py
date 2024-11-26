from module_moyenne import calcul_moyennes_notes_UE1_S1
from module_moyenne import calcul_moyennes_notes_UE2_S1
from module_moyenne import calcul_moyennes_notes_UE3_S1
from module_moyenne import generate_html_file

# Appel de la fonction pour calculer les moyennes et générer le fichier HTML
donnees_UE1, donnees_UE1_S2, resultat, data_rows_UE1 = calcul_moyennes_notes_UE1_S1()
donnees_UE2, donnees_UE2_S2, resultatUE2, data_rowsUE2 = calcul_moyennes_notes_UE2_S1()
donnees_UE3, donnees_UE3_S2, resultatUE3, data_rowsUE3 = calcul_moyennes_notes_UE3_S1()
generate_html_file("moyenne_UE1.html", "Moyenne UE1", ["Nom", "Prenom", "Moyenne UE1 S1", "Moyenne UE1 S2", "Resultat", ], ["Nom", "Prenom", "Moyenne UE2 S1", "Moyenne UE2 S2", "Resultat UE2"],["Nom", "Prenom", "Moyenne UE3 S1", "Moyenne UE3 S2", "Resultat UE3"],data_rows_UE1, data_rowsUE2, data_rowsUE3)
