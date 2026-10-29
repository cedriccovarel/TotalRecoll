# TotalRecoll

Recolle le tableau « Tous contrats » (Salesforce) avec les données techniques d’ExtracTerre pour produire le tableau d’opérations de l’Observatoire (1 ligne = 1 bâtiment).

- Entrées : « Tous contrats » (.xlsx, colonnes « Numéro du contrat » et « Opération: Code interne ») ; une ou plusieurs sources ExtracTerre (session .json ou export Multi-projets .xlsx, feuille « Données par bâtiment »).
- Clés : numéro de contrat (format 0000xxxx) et/ou numéro d’opération (OPE-XXXXXX). Un contrat et une OPE qui désignent deux opérations différentes sont signalés en conflit, jamais fusionnés.
- Colonnes administratives : Salesforce fait foi. Colonnes techniques : ExtracTerre, sauf option « garder la valeur de Tous contrats » ; chaque écart est listé (onglet « Écarts » de l’export).
- Contrôle : taux de recollement, lignes à contrôler (conflits, opérations absentes de Tous contrats, découpage en bâtiments à vérifier), écarts tranchés un par un (choix gardés dans ce navigateur), carte des bâtiments géolocalisés par ExtracTerre (fond IGN), correspondance des colonnes, fiche détaillée par ligne.
- Export : copie pour Google Sheets (coller en A1) ou .xlsx (OPERATIONS, Écarts, À contrôler ; Google Sheets → Fichier → Importer).
- Tout est traité dans le navigateur ; aucun fichier client n’est envoyé ni versionné.

`fields.js` est généré depuis `js/config.js` d’ExtracTerre (FIELD_DEFS, familles techniques).

DH (ExtracTerre ≥ 2.11) : « DH projet » reçoit le plus défavorable de DH ZT / DH ZNT ; un DH sans type de zone (exports plus anciens) est reporté en DH ZNT, ou en DH ZT si le bâtiment n’a que des logements traversants. Les colonnes DH ZT, DH ZNT, ITE, ITI, Brasseur d’air et Climatisation active sont ajoutées en fin de tableau quand OPERATIONS ne les a pas.

Tests : `node test_recoll.mjs` (données fictives).
