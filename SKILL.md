---
name: outlook-entra
description: |
  Microsoft Outlook via OAuth 2.0 (device code flow) et Microsoft Graph API.
  Lecture seule — Mail.Read, Calendars.Read, Contacts.Read.
metadata:
  author: custom
  version: "1.0"
---

# Outlook Entra — SKILL.md

Microsoft Outlook via OAuth 2.0 (device code flow) et Microsoft Graph API.
**Lecture seule** — seules les permissions `Mail.Read`, `Calendars.Read`, `Contacts.Read` sont utilisées.

## Prérequis

- App enregistrée sur **Entra** (Azure AD) avec permissions :
  - `Mail.Read`, `Calendars.Read`, `Contacts.Read`
  - OAuth 2.0 device code flow activé
- Python 3.8+ avec `requests`
- Appairage d'un **nœud OpenClaw** avec capability `http` (requis uniquement si l'IP du serveur est blacklisted par Microsoft)
- Fichier `.env` configuré (voir `.env.example`)

## Installation

```bash
# Dépendance Python
uv pip install requests cryptography html2text

# Copier et éditer la config
cp .env.example .env
# ⚠️ Remplir client_id, client_secret, oauth urls dans .env
```

## Authentification (Flow)

Le device code flow nécessite de contacter Microsoft. Si l'IP du serveur est blacklistée (souvent le cas), **la première requête (device code) passe toujours par le nœud Android**. Les requêtes suivantes (échange de tokens, appels Graph API) sont exécutées en local par les scripts Python.

---

### Étape 1 — Demander un device code (via nœud Android)

```bash
AZURE_CLIENT_ID=$(grep AZURE_CLIENT_ID .env | cut -d= -f2)
AZURE_TENANT_ID=$(grep AZURE_TENANT_ID .env | cut -d= -f2)
SCOPES_DEVICE_CODE=$(grep SCOPES_DEVICE_CODE .env | cut -d= -f2 | tr -d '"')

openclaw nodes invoke \
  --node "S25+ de Frederic" \
  --command "http.request" \
  --params "{\"url\": \"https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/devicecode\", \"method\": \"POST\", \"headers\": {\"Content-Type\": \"application/x-www-form-urlencoded\"}, \"body\": \"client_id=${AZURE_CLIENT_ID}&scope=${SCOPES_DEVICE_CODE}\"}"
```

Réponse : `user_code` (ex: `JE2ZZYDA7`) + `device_code`.

---

### Étape 2 — L'utilisateur s'authentifie

Aller sur **https://login.microsoft.com/device** et entrer le `user_code`. Délai : 15 minutes.

---

### Étape 3 — Lancer le script d'auth (échange device_code → tokens)

Une fois l'authentification faite, lancer le script Python pour échanger le device_code et sauvegarder les tokens :

```bash
python scripts/outlook_auth.py
```

Le script va automatiquement détecter qu'il n'y a pas de token valide et lancer le polling du device_code (en utilisant les variables du .env).

> Si le serveur ne peut pas contacter Microsoft directement, le polling échouera. Voir "Dépannage" ci-dessous.

---

### Commandes utiles

```bash
# Vérifier le statut du token
python scripts/outlook_auth.py --status

# Révoquer et supprimer les tokens
python scripts/outlook_auth.py --revoke
```

## Commandes (lecture seule)

```bash
# Statut de connexion
python scripts/outlook_auth.py --status

# Lire les derniers messages
python scripts/outlook_graph.py messages --folder Inbox --top 10

# Détail d'un message (corps complet, Markdown par défaut)
python scripts/outlook_graph.py message <messageId>

# Détail en HTML brut (pour extraction/collage)
python scripts/outlook_graph.py message <messageId> --raw

# Lister les dossiers mail
python scripts/outlook_graph.py folders

# Pièces jointes d'un message
python scripts/outlook_graph.py attachments <messageId>

# Télécharger une pièce jointe
python scripts/outlook_graph.py download <attachmentId>

# Événements calendrier
python scripts/outlook_graph.py events --top 10

# Contacts
python scripts/outlook_graph.py contacts --top 20

# Rechercher dans les mails
python scripts/outlook_graph.py search "mot-clé"

# Profil utilisateur
python scripts/outlook_graph.py profile
```

## Variables d'environnement (.env)

| Variable | Description | Exemple |
|---|---|---|
| `AZURE_TENANT_ID` | GUID du tenant Entra | `52ffb8b9-…` |
| `AZURE_CLIENT_ID` | ID de l'app (Application ID) | `xxxxxxxx-…` |
| `AZURE_CLIENT_SECRET` | Secret de l'app | `~` |
| `AZURE_REDIRECT_URI` | Redirect URI (device flow : tout fait) | `http://localhost` |
| `OAUTH_TOKEN_URL` | URL token endpoint | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token` |
| `OAUTH_DEVICE_CODE_URL` | URL device code endpoint | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode` |
| `MS_GRAPH_BASE_URL` | Base URL Microsoft Graph | `https://graph.microsoft.com/v1.0` |
| `TOKEN_FILE` | Chemin du fichier de stockage des tokens | `~/.openclaw/outlook_tokens.json` |
| `TOKEN_FILE_KEY` | Clé de chiffrement (optionnel) | _(vide par défaut)_ |

## Structure du skill

```
outlook-entra/
├── SKILL.md
├── README.md
├── .env.example
├── .gitignore
├── scripts/
│   ├── outlook_auth.py      # OAuth device code flow + refresh
│   ├── outlook_graph.py     # Appels Graph API (lecture seule)
│   └── outlook_token.py     # Module partagé (lecture/refresh tokens)
└── tests/
    └── test_outlook.py      # Tests unitaires
```

## Notes

- Le **device code flow** (RFC 8628) : l'utilisateur authentifie via `https://microsoft.com/devicelogin`. Une seule fois.
- Les **refresh tokens** sont automatiquement utilisés quand l'access token expire.
- Si `TOKEN_FILE_KEY` est défini, les tokens sont chiffrés AES-GCM avant stockage.
- Les erreurs 401 du Graph API déclenchent un refresh automatique.

## Ressources

- [OAuth 2.0 Device Code Flow — Microsoft](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Microsoft Graph — Mail API](https://learn.microsoft.com/en-us/graph/api/user-list-messages)
- [Microsoft Graph — Calendar API](https://learn.microsoft.com/en-us/graph/api/user-list-events)
