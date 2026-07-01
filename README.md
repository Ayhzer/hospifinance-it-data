# hospifinance-it-data

Jeu de **données de démonstration** pour l'application Hospifinance-IT.

## Contenu

| Fichier | Description |
|---------|-------------|
| `data/users.json` | Comptes utilisateurs de démo (identifiants, rôles, mots de passe encodés en base64) |
| `data/opex.json` | Lignes budgétaires OPEX (démo) |
| `data/opex-orders.json` | Commandes OPEX (démo) |
| `data/capex.json` | Projets d'investissement CAPEX (démo) |
| `data/capex-orders.json` | Commandes CAPEX (démo) |
| `data/eprd.json` | Budgets EPRD par compte (démo) |
| `data/reclassement.json` | Nomenclature analytique + règles de reclassement (démo) |
| `data/settings.json` | Paramètres applicatifs (couleurs, colonnes, préférences, source d'import automatique) |

## Utilisation

Ce dépôt sert de backend de données à Hospifinance-IT, soit :
- en **local** via `local-server.js` (lecture/écriture des fichiers `data/*.json`),
- en **production** via la synchronisation GitHub intégrée (token configuré dans les paramètres de l'application).

## Production : données réelles

Les données ci-dessus sont **fictives** et servent uniquement de démonstration.
Pour exploiter des données réelles d'établissement :

1. Créez un dépôt **privé** dédié (ne jamais exposer publiquement des données réelles).
2. Renseignez `VITE_GITHUB_OWNER` / `VITE_GITHUB_REPO` (cf. `.env.example` de l'app).
3. Importez vos commandes via le modèle canonique (fichier exemple téléchargeable dans l'app).
