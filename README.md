# outlook-entra

Microsoft Outlook / Office 365 via OAuth 2.0 (device code flow) et Microsoft Graph API.
**Lecture seule** — permissions `Mail.Read`, `Calendars.Read`, `Contacts.Read`.

## Ce que fait ce skill

- **Lire** des emails, événements calendrier, contacts et pièces jointes
- **Rechercher** dans les emails
- Token OAuth chiffré (AES-GCM) avant stockage

## Prérequis

- Python 3.8+ (`requests`, `cryptography`)
- App enregistrée sur **Microsoft Entra** avec permissions :
  - `Mail.Read`, `Calendars.Read`, `Contacts.Read`
  - OAuth 2.0 device code flow activé
- Appairage d'un **nœud OpenClaw** avec capability `http` (requis uniquement si l'IP du serveur est blacklisted par Microsoft)

## Installation

```bash
# 1. Cloner le repo
git clone https://github.com/fredguile/openclaw-skill-outlook-entra.git
cd openclaw-skill-outlook-entra

# 2. Créer et remplir la config
cp .env.example .env
# Éditer .env : AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET

# 3. Installer les dépendances Python
uv pip install requests cryptography
```

## Première authentification

Si l'IP du serveur est blacklisted par Microsoft (requiert un nœud avec `http`) :

```bash
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

Entrer le `user_code` sur **https://microsoft.com/devicelogin**

Puis échanger le code contre un token :

```bash
curl -X POST \
  -d "grant_type=urn:ietf:params:oauth:grant-type:device_code" \
  -d "client_id=<CLIENT_ID>" \
  -d "device_code=<DEVICE_CODE>" \
  "https://login.microsoftonline.com/<TENANT>/oauth2/v2.0/token"
```

Alternativement, utiliser le script automatique (si le serveur n'est pas blacklisted) :

```bash
python scripts/outlook_auth.py
```

## Commandes

```bash
# Statut du token
python scripts/outlook_auth.py --status

# Lire les derniers messages
python scripts/outlook_graph.py messages --folder Inbox --top 10

# Détail d'un message (corps complet)
python scripts/outlook_graph.py message <messageId>

# Pièces jointes d'un message
python scripts/outlook_graph.py attachments <messageId>

# Télécharger une pièce jointe
python scripts/outlook_graph.py download <messageId> --attach-id <id>

# Dossiers mail
python scripts/outlook_graph.py folders

# Calendrier
python scripts/outlook_graph.py events --top 10

# Contacts
python scripts/outlook_graph.py contacts --top 20

# Rechercher dans les mails
python scripts/outlook_graph.py search "mot-clé"

# Profil utilisateur
python scripts/outlook_graph.py profile

# Supprimer le token
python scripts/outlook_auth.py --revoke
```

## Variables d'environnement (.env)

| Variable | Description |
|---|---|
| `AZURE_TENANT_ID` | GUID du tenant Entra |
| `AZURE_CLIENT_ID` | Application (client) ID |
| `AZURE_CLIENT_SECRET` | Client secret |
| `OAUTH_DEVICE_CODE_URL` | Device code endpoint |
| `OAUTH_TOKEN_URL` | Token endpoint |
| `MS_GRAPH_BASE_URL` | Base URL Graph API |
| `TOKEN_FILE` | Chemin du fichier de tokens |
| `TOKEN_FILE_KEY` | Clé de chiffrement (optionnel) |
| `SCOPES_DEVICE_CODE` | Scopes OAuth demandés |
| `REQUEST_TIMEOUT` | Timeout requêtes HTTP |

## Chiffrement des tokens

Les tokens sont chiffrés AES-GCM si `TOKEN_FILE_KEY` est défini dans `.env` :

```bash
openssl rand -base64 32
```

## Ressources

- [OAuth 2.0 Device Code Flow — Microsoft](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Microsoft Graph — Mail API](https://learn.microsoft.com/en-us/graph/api/user-list-messages)
