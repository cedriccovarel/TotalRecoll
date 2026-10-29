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

## ANRU (NPNRU)

Colonnes **ANRU** (Oui / Non) et **Quartier ANRU** ajoutées avant « Recollement », et onglet **ANRU** : carte des bâtiments situés dans un quartier du NPNRU, filtrée par les mentions de l’opération (« Opération: Mentions », cases à cocher).

- Quartiers : les 480 quartiers du NPNRU, 216 d’intérêt national (arrêté du 29 avril 2015) et 264 d’intérêt régional (arrêté du 15 janvier 2019), liste ANCT / ANRU publiée sur data.gouv.fr (COG 2022), avec les périmètres des quartiers prioritaires 2015 (ANCT, data.gouv.fr). Licence Ouverte 2.0. Fichier `data/npnru.json` (contours simplifiés à environ 2 m).
- Bâtiment géolocalisé par ExtracTerre (Latitude / Longitude) : Oui s’il est dans un périmètre ou à 10 m au plus de son bord (distance indiquée dans « Quartier ANRU »), Non sinon.
- Carte (fond IGN sans filtre) : quartiers d’intérêt national en rouille, d’intérêt régional en bleu pointillé ; chaque bâtiment prend la couleur de son quartier.
- « Exporter en CSV » : les bâtiments affichés (mentions cochées), toutes colonnes, séparateur « ; », UTF-8 avec BOM (Excel).
- Sans coordonnées : Non si la commune (code postal + ville) n’a aucun quartier NPNRU ; vide (« à vérifier ») si elle en a un.
- Périmètre strict : la bande de 300 m autour des quartiers n’est pas prise en compte.
