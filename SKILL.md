# Outlook Entra — SKILL.md

Microsoft Outlook via OAuth 2.0 (device code flow) et Microsoft Graph API.
Utilise l'app Entra enregistrée de l'utilisateur — aucun service tiers.

## Prérequis

- App enregistrée sur **Entra** (Azure AD) avec permissions :
  - `Mail.Read`, `Mail.Send`, `Calendars.Read`, `Contacts.Read`
  - OAuth 2.0 device code flow activé
- Python 3.8+ avec `requests`
- Appairage d'un **nœud OpenClaw** avec capability `http` (ex: smartphone Android)
  — utilisé pour la requête `/devicecode` car l'IP du serveur est blacklistée par Microsoft
- Fichier `.env` configuré (voir `.env.example`)

## Installation

```bash
# Dépendance Python
uv pip install requests

# Copier et éditer la config
cp .env.example .env
# ⚠️ Remplir client_id, client_secret, oauth urls dans .env
```

## Authentification

### Première authentification

Important: étant donné que Microsoft rejette les requêtes `/devicecode` depuis le serveur, il faut IMPERATIVEMENT utiliser le nœud HTTP pour faire cette première requête:

```bash
# Via le nœud OpenClaw (ex: S25+ de Frederic)
openclaw nodes invoke \
  --node "S25+ de Frederic" \
  --command "http.request" \
  --params '{
    "url": "https://login.microsoftonline.com/<TENANT>/oauth2/v2.0/devicecode",
    "method": "POST",
    "headers": {"Content-Type": "application/x-www-form-urlencoded"},
    "body": "client_id=<CLIENT_ID>&scope=user.read%20openid%20profile%20offline_access"
  }'
```

**Le `<TENANT>` dépend de l'organisation :**

- Organisations standard : GUID du tenant (ex: `52ffb8b9-c339-49f6-97ba-7c9bb2ff7482`)

→ Le résultat contient `user_code` et `device_code`. Entrer le code sur **https://microsoft.com/devicelogin**

Puis récupérer le token via poll (à faire directement depuis le serveur ou via le nœud) :

```bash
curl -X POST \
  -d "grant_type=urn:ietf:params:oauth:grant-type:device_code" \
  -d "client_id=<CLIENT_ID>" \
  -d "device_code=<DEVICE_CODE>" \
  "https://login.microsoftonline.com/<TENANT>/oauth2/v2.0/token"
```

### Script automatique (si le serveur n'est pas blacklisted)

```bash
python scripts/outlook_auth.py              # Lance le device code flow
python scripts/outlook_auth.py --status     # Statut du token
python scripts/outlook_auth.py --revoke     # Révoque et supprime le token
```

## Commandes principales

```bash
# Statut de connexion
python scripts/outlook_auth.py --status

# Lire les derniers messages
python scripts/outlook_graph.py messages --folder Inbox --top 10

# Lire un message par ID
python scripts/outlook_graph.py message <messageId>

# Envoyer un message
python scripts/outlook_graph.py send \
  --to recipient@example.com \
  --subject "Sujet" \
  --body "Corps du message"

# Lister les dossiers mail
python scripts/outlook_graph.py folders

# Événements calendrier
python scripts/outlook_graph.py events --top 10

# Contacts
python scripts/outlook_graph.py contacts --top 20

# Marquer message lu / non lu
python scripts/outlook_graph.py mark-read <messageId>
python scripts/outlook_graph.py mark-unread <messageId>

# Rechercher dans les mails
python scripts/outlook_graph.py search "mot-clé"
```

## Variables d'environnement (.env)

| Variable                | Description                              | Exemple                                                             |
| ----------------------- | ---------------------------------------- | ------------------------------------------------------------------- |
| `AZURE_TENANT_ID`       | ID du tenant Entra ou nom de domaine     | `fairleaonline.com`                                                 |
| `AZURE_CLIENT_ID`       | ID de l'app (Application ID)             | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`                              |
| `AZURE_CLIENT_SECRET`   | Secret de l'app                          | `~`                                                                 |
| `OAUTH_TOKEN_URL`       | URL token endpoint                       | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`      |
| `OAUTH_DEVICE_CODE_URL` | URL device code endpoint                 | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode` |
| `MS_GRAPH_BASE_URL`     | Base URL Microsoft Graph                 | `https://graph.microsoft.com/v1.0`                                  |
| `TOKEN_FILE`            | Chemin du fichier de stockage des tokens | `~/.openclaw/outlook_tokens.json`                                   |
| `TOKEN_FILE_KEY`        | Clé de chiffrement (optionnel)           | _(vide par défaut)_                                                 |

## Structure du skill

```
outlook-entra/
├── SKILL.md
├── .env.example
├── scripts/
│   ├── outlook_auth.py      # OAuth device code flow + refresh
│   ├── outlook_graph.py     # Appels Graph API (mail, calendar, contacts)
│   └── outlook_token.py     # Module partagé (lecture/refresh tokens)
└── tests/
    └── test_outlook.py      # Tests unitaires
```

## Notes

- Le **device code flow** (RFC 8628) : l'utilisateur authentifie via `https://microsoft.com/devicelogin`. Une seule fois.
- Les **refresh tokens** sont automatiquement utilisés quand l'access token expire.
- Les **subscriptions** (webhooks) nécessitent un endpoint HTTPS public — non supporté. Utilisez un cron pour le polling.
- Si `TOKEN_FILE_KEY` est défini, les tokens sont chiffrés AES-GCM avant stockage.
- Les erreurs 401 du Graph API déclenchent un refresh automatique.

## Ressources

- [OAuth 2.0 Device Code Flow — Microsoft](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Microsoft Graph — Mail API](https://learn.microsoft.com/en-us/graph/api/user-list-messages)
- [Microsoft Graph — Calendar API](https://learn.microsoft.com/en-us/graph/api/user-list-events)
