# TotalRecoll

Recolle le tableau « Tous contrats » (Salesforce) avec les données techniques d’ExtracTerre pour produire le tableau d’opérations de l’Observatoire (1 ligne = 1 bâtiment).

- Entrées : « Tous contrats » (.xlsx, colonnes « Numéro du contrat » et « Opération: Code interne ») ; une ou plusieurs sources ExtracTerre (session .json ou export Multi-projets .xlsx, feuille « Données par bâtiment »).
- Clés : numéro de contrat (format 0000xxxx) et/ou numéro d’opération (OPE-XXXXXX). Sans colonne contrat, le numéro est lu en tête du nom de l’opération ou du dossier (« 00004701 - 26 - … »). Un contrat et une OPE qui désignent deux opérations différentes sont signalés en conflit, jamais fusionnés.
- Colonnes administratives : Salesforce fait foi. Colonnes techniques : ExtracTerre, sauf option « garder la valeur de Tous contrats » ; chaque écart est listé (onglet « Écarts » de l’export).
- Contrôle : taux de recollement, lignes à contrôler (conflits, opérations absentes de Tous contrats, découpage en bâtiments à vérifier), écarts tranchés un par un (choix gardés dans ce navigateur), carte des bâtiments géolocalisés par ExtracTerre (fond IGN), correspondance des colonnes, fiche détaillée par ligne.
- Envoi à l’Observatoire : ouvrir TotalRecoll depuis l’Observatoire (fenêtre Données › Complément TotalRecoll), faire le recollement, puis « Envoyer à l’Observatoire ». Seules les valeurs écrites par ExtracTerre partent, d’onglet à onglet (même origine) ; l’Observatoire remplit les cellules vides (ou bouche-trous « NC », « - »…) de son Sheet OPERATIONS, corrige les codes postaux / INSEE / coordonnées / années illisibles, ajoute à l’opération les bâtiments que le Sheet n’a pas, et répond par un bilan (cellules remplies, écarts, opérations absentes, bâtiments ajoutés, valeurs corrigées). Rien n’est stocké ni envoyé à un serveur.
- Export : copie pour Google Sheets (coller en A1) ou .xlsx (OPERATIONS, Écarts, À contrôler ; Google Sheets → Fichier → Importer).
- Tout est traité dans le navigateur ; aucun fichier client n’est envoyé ni versionné.

`fields.js` est généré depuis `js/config.js` d’ExtracTerre (FIELD_DEFS, familles techniques).

DH (ExtracTerre ≥ 2.11) : « DH projet » reçoit le plus défavorable de DH ZT / DH ZNT ; un DH sans type de zone (exports plus anciens) est reporté en DH ZNT, ou en DH ZT si le bâtiment n’a que des logements traversants. Les colonnes DH ZT, DH ZNT, ITE, ITI, Brasseur d’air et Climatisation active sont ajoutées en fin de tableau quand OPERATIONS ne les a pas.

Tests : `node test_recoll.mjs` (données fictives).

## Confort d’été

Onglet **Confort d’été** : par bâtiment, DH des zones traversantes et non traversantes, système constructif, brasseur d’air (Oui / Non), bâtiment climatisé (Oui / Non), exportable en CSV. Valeurs lues dans les colonnes recollées (Tous contrats + ExtracTerre) ; rien n’est ajouté à l’export OPERATIONS.

- Système constructif : « Ossature bois » (dont CLT, bois massif) ; « Béton ou brique + isolation extérieure » si ITE ; « Béton ou brique + isolation intérieure ou isolation répartie » si ITI, monomur ou béton cellulaire.
- Une structure ou une isolation absente, mixte (ITE et ITI) ou hors catégories reste vide ; la raison est dans « À vérifier ».

## ANRU (NPNRU)

Colonnes **ANRU** (Oui / Non) et **Quartier ANRU** ajoutées avant « Recollement », et onglet **ANRU** : carte des bâtiments situés dans un quartier du NPNRU, filtrée par les mentions de l’opération (« Opération: Mentions », cases à cocher).

- Quartiers : les 480 quartiers du NPNRU, 216 d’intérêt national (arrêté du 29 avril 2015) et 264 d’intérêt régional (arrêté du 15 janvier 2019), liste ANCT / ANRU publiée sur data.gouv.fr (COG 2022), avec les périmètres des quartiers prioritaires 2015 (ANCT, data.gouv.fr). Licence Ouverte 2.0. Fichier `data/npnru.json` (contours simplifiés à environ 2 m).
- Bâtiment géolocalisé par ExtracTerre (Latitude / Longitude) : Oui s’il est dans un périmètre ou à 10 m au plus de son bord (distance indiquée dans « Quartier ANRU »), Non sinon.
- Carte (fond IGN sans filtre) : quartiers d’intérêt national en rouille, d’intérêt régional en bleu pointillé ; chaque bâtiment prend la couleur de son quartier.
- « Exporter en CSV » : les bâtiments affichés (mentions cochées), toutes colonnes, séparateur « ; », UTF-8 avec BOM (Excel).
- Sans coordonnées : Non si la commune (code postal + ville) n’a aucun quartier NPNRU ; vide (« à vérifier ») si elle en a un.
- Périmètre strict : la bande de 300 m autour des quartiers n’est pas prise en compte.

## Vérification des « Oui » (onglet Recherches)

Bouton **✔ Vérifier les « Oui »** : chaque réponse Oui / Partiel d'un critère de la grille choisie (sélection de l'onglet
Recherches) est montrée avec son **exigence** et sa **vérification** (document, page, ligne lue et contexte, pour contrôler à la
main dans la source). Deux affichages : **Tableau** (boutons Accepter / Rejeter) ou **Cartes** à glisser (droite = Oui,
gauche = Non, vers le haut = Partiel, vers le bas = Sans objet ; clavier → ← ↑ ↓ ; Retour arrière = annuler). Filtres par critère et par statut (à vérifier, acceptés, rejetés).
Un « Non » compte comme indice contraire (non conforme à l'exigence) ; « Sans objet » (la phrase n'a pas de lien avec l'exigence)
compte comme nul (ni trouvé, ni contraire), noté à part pour corriger les règles ; une réponse « Partiel » est gardée comme partielle (comme dans ExtracTerre, elle
ne compte pas comme trouvée pour le thème) : statistiques, thèmes (recalculés) et CSV. Un Oui donné une fois dans une opération vaut pour les cartes en attente de la même opération qui portent sur le même
critère ou montrent la même phrase (ex. 44.2 des fenêtres et du hall) : elles ne sont plus redemandées. Les décisions sont gardées dans le
navigateur (bouton « Tout effacer » pour recommencer ; celles prises avant les règles du 18/11 — Non = contraire, Sans objet = nul —
ont été effacées) et s'exportent / s'importent en fichier JSON (sauvegarde, partage). Le contexte des preuves vient des sessions
exportées par ExtracTerre 3.0.4 ou plus récent.
