# TotalRecoll

Recolle le tableau « Tous contrats » (Salesforce) avec les données techniques d’ExtracTerre pour produire le tableau d’opérations de l’Observatoire (1 ligne = 1 bâtiment).

- Entrées : « Tous contrats » (.xlsx, colonnes « Numéro du contrat » et « Opération: Code interne ») ; une ou plusieurs sources ExtracTerre (session .json ou export Multi-projets .xlsx, feuille « Données par bâtiment »).
- Clés : numéro de contrat (format 0000xxxx) et/ou numéro d’opération (OPE-XXXXXX). Un contrat et une OPE qui désignent deux opérations différentes sont signalés en conflit, jamais fusionnés.
- Colonnes administratives : Salesforce fait foi. Colonnes techniques : ExtracTerre, sauf option « garder la valeur de Tous contrats » ; chaque écart est listé (onglet « Écarts » de l’export).
- Export : copie pour Google Sheets (coller en A1) ou .xlsx (Google Sheets → Fichier → Importer).
- Tout est traité dans le navigateur ; aucun fichier client n’est envoyé ni versionné.

`fields.js` est généré depuis `js/config.js` d’ExtracTerre (FIELD_DEFS, familles techniques).

Tests : `node test_recoll.mjs` (données fictives).
