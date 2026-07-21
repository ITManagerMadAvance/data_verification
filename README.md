# data_verification

Vérification automatisée des données mWater pour l'activité **Appel maintenance préventive**, sur la base des six dimensions du Manuel de vérification de données MadAvance (Complétude, Promptitude, Validité, Unicité, Cohérence, Fiabilité).

Le script `verify_maintenance_preventive.py` :

1. S'authentifie sur l'API mWater et télécharge 4 datagrids déjà configurés dans le portail (Appel maintenance préventive, Maintenance préventive, Réparation après panne, Première réhabilitation).
2. Applique les règles de vérification (voir le détail dans le manuel ClickUp lié).
3. Génère un rapport Excel horodaté `data_verification_call_JJMMAAAA.xlsx` (onglet détail + onglet résumé par dimension).
4. Dépose le rapport sur SharePoint via Microsoft Graph.
5. Envoie un email de confirmation avec le décompte d'anomalies par dimension.

Exécution automatique hebdomadaire via GitHub Actions (lundi 06:30 UTC), ou manuelle via `workflow_dispatch`. Ne modifie jamais les données mWater — uniquement de la détection/reporting. La correction reste une étape manuelle séparée.

## Secrets GitHub Actions requis (Settings > Secrets and variables > Actions)

| Secret | Description |
| --- | --- |
| `MWATER_USERNAME` / `MWATER_PASSWORD` | Identifiants mWater |
| `AZURE_TENANT_ID` | Même App Registration que le repo `mWater_backup` ("BackupOffice365") |
| `AZURE_CLIENT_ID` | idem |
| `AZURE_CLIENT_SECRET` | idem — généré dans Entra ID, valeur connue uniquement dans GitHub Secrets |
| `SHAREPOINT_DRIVE_ID` | Cible définitive : dossier [Vérification des données](https://madavancengo.sharepoint.com/:f:/s/ITMadAvance/IgAwR3jEg9UMT7JcKHJAeQttAXCrIid1qcCeAEtuxAPEdNU). **Valeur temporaire de test** (même site `ITMadAvance` que le backup) : `b!8D4xOy74F0-I2pDx1b5rX8HkGdhgNxpGpD3JvyEKMY4-rXDAT44VQ40NYtVFZG-V` |
| `SHAREPOINT_FOLDER_ITEM_ID` | Cible définitive : à résoudre (voir ci-dessous). **Valeur temporaire de test** : `01T5F36LFRTB6Z2I6KHNHLORRAGKD554ZL` (dossier "Backup" du repo `mWater_backup` — le rapport de test y atterrira, à déplacer/nettoyer ensuite) |
| `EMAIL_SENDER` | Boîte d'envoi (ex. `it@madavance.org`) |
| `EMAIL_RECIPIENTS` | Destinataire(s), séparés par des virgules |

Ces secrets ne sont pas partagés automatiquement entre repos GitHub : même si `AZURE_TENANT_ID`/`AZURE_CLIENT_ID`/`AZURE_CLIENT_SECRET` existent déjà dans `mWater_backup`, il faut les recopier dans les secrets de **ce** repo.

### Résoudre `SHAREPOINT_DRIVE_ID` / `SHAREPOINT_FOLDER_ITEM_ID`

Le dossier cible a été partagé via un lien de partage, pas un chemin direct. Comme pour le backup, l'adressage par chemin (`/sites/{id}/drive/root:/{chemin}`) est cassé sur ce tenant (`400 Resource not found for the segment 'root:'`) — il faut résoudre le lien en `driveId` + ID d'item via l'API Graph `/shares/{shareId}/driveItem`, ou via Graph Explorer par navigation. Non résolu à ce stade (nécessite une session Microsoft 365 authentifiée) — à compléter avant la première exécution.

## Datagrids mWater utilisés (IDs codés en dur dans le script, pas des secrets)

| Activité | Formulaire | Datagrid |
| --- | --- | --- |
| Appel maintenance préventive | `c08b3fe26d0f42c084074701f29eb75e` | `5f443b8a8c144502a304c8c5c24d4f82` |
| Maintenance préventive | `de26d89a5c8a4452b42158c622be20d0` | `2cfe0ba7ac264d119cfc8964b5f3cebc` |
| Réparation après panne | `958b4763788348d699e7d8c5821f92ee` | `d9f1c36a2d6340429658b6628fe81b88` |
| Première réhabilitation (datagrid mixte, filtré dans le script) | `86cf66efdd3749dd8a121314bab3675a` | `0638d32971704164b9ac22d549d4818e` |

## Règles appliquées

- **Complétude** : `Status = Draft` (brouillon jamais soumis) ou `Status = Final` sans Water Point ID renseigné.
- **Promptitude** : `Drafted On <= Submitted On`.
- **Validité** : `Signal reference` doit respecter `{DEPLOYMENT}_{JJMMAAAA}_{E|S}{N}` ; Water Point ID doit être numérique.
- **Unicité** : doublons de `Signal reference`, ou rejet mWater contenant "doublon" dans `Rejection message`.
- **Cohérence** : préfixe de déploiement cohérent ; si pompe "Partially" → code présent dans Maintenance préventive ; si "No" → présent dans Réparation après panne ; date embarquée dans le code ≤ `Completion date of the work` du formulaire correspondant.
- **Fiabilité** : rejets contenant "ID incorrect" croisés avec les Water Point ID des Premières réhabilitations réussies (`Type de travaux = "Première réhabilitation"` et `Status = Final`) dans le datagrid Réhabilitation ; si le WP ID est trouvé ailleurs dans ce même datagrid (autre type/statut), le contexte est indiqué pour réexamen.

## Test en local

```bash
pip install -r requirements.txt
export MWATER_USERNAME=... MWATER_PASSWORD=...
export AZURE_TENANT_ID=... AZURE_CLIENT_ID=... AZURE_CLIENT_SECRET=...
export SHAREPOINT_DRIVE_ID=... SHAREPOINT_FOLDER_ITEM_ID=...
export EMAIL_SENDER=... EMAIL_RECIPIENTS=...
python verify_maintenance_preventive.py
```

## Test de validation (logique seule, sans upload)

Testé le 21/07/2026 sur un export réel des 4 datagrids (3073 réponses Appel, 732 Maintenance, 171 Réparation, 1688 Réhabilitation) : 313 anomalies détectées (Cohérence 186, Unicité 57, Complétude 40, Promptitude 29, Fiabilité 1). Le nombre de cas Cohérence est nettement supérieur aux quelques exemples relevés manuellement dans le manuel de vérification — probablement parce que la vérification manuelle n'avait couvert qu'un échantillon, pas l'historique complet. **À valider avec Lanja avant mise en production** : passer en revue un échantillon des anomalies "absent du formulaire" pour confirmer qu'il ne s'agit pas de faux positifs (ex. suivi encore en attente plutôt qu'anomalie réelle).
