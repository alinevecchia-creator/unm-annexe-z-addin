# Annexe Z – Add-in Word v3 (Test UNM Machines)

Cette version **test** fournit :
- Panneau latéral (Exigences RM / Phénomènes dangereux)
- Stockage local (JSON caché dans le document)
- Génération Annexe Z
- **Simulation** de synchronisation SharePoint (pas d’authentification fournie dans cette version)

## Démarrage (local)
1. Servez ce dossier via HTTPS (ex. `npx http-server -S -C cert.pem -K key.pem` ou `office-addin-debugging`).
2. Word → Insertion → Mes compléments → Charger un complément personnalisé → `manifest.xml`.
3. Ouvrez votre modèle Word et placez `[ANNEXE_Z_TABLE]` à l’endroit d’insertion de la table.

## Configuration SharePoint
- Éditez `config.json` (tenant, siteUrl, noms de listes).
- La v3.1 ajoutera l’authentification Graph (SSO) et les appels réels d’upsert.

## Listes SharePoint (schéma cible)
- `Projects` : ProjectId (clé), TC, Titre, Statut, Période.
- `RequirementsCoverage` : ExigenceId, ProjectId (lookup), ClauseNumber, ClauseStableId, LastUpdateUtc.
- `Hazards` : ProjectId (lookup), HazardId, HazardLabel, Category, LastUpdateUtc.

## Script d’initialisation
Voir `pnp-init.ps1` pour créer listes/colonnes/index/vues via **PnP.PowerShell**.
