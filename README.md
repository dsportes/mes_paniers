# Exécution

Export XLS depuis Odoo : Points de Vente >>> Commandes >>> Commandes

Un filtre sur dates est possible.

Sélectionner tout.

Export avec les colonnes : 
- 'Client/Nom affiché'
- 'Date de la commande'
- 'Total'

Garder le nom de fichier proposé : pos.order.xls et le sauver sur le bureau.

Exécution dans le directory de l'application "paniers" :

node main.js

Des arguments facultatifs peuvent être donnés après "main.js":
- entree=mon_fichier1 : path du fichier d'entrée, par défaut c'est ~/Desktop/pos.order.xls
- entree=mon_fichier2 : path du fichier de sortie, par défaut c'est ~/Desktop/paniers.xlsx
- nbmoisactif=3 : un coop est considéré comme actif pour un mois M s'il a fait au moins un achat au cours des trois derniers mois (M inclus)
- nbtranches=5 : les moyennes des paniers sont décomptés sur 5 tranches (vingtiles), ou 10 tranches (déciles)
- voirnoms=vrai : par défaut les ventes des coops sur les mois considérés sont masqués mais on peut forcer à les voir.
