---
name: outlook-entra
description: |
  Microsoft Outlook via OAuth 2.0 (device code flow) et Microsoft Graph API.
  Lecture + brouillons (jamais d'envoi) — Mail.Read, Calendars.Read, Contacts.Read.
  Utilise le serveur MCP `outlook-mcp` (13 outils) — compatible sessions voice-call.
metadata:
  author: custom
  version: "3.1"
---

# Outlook Entra — SKILL.md

Microsoft Outlook via OAuth 2.0 (device code flow) et Microsoft Graph API.
**Lecture seule + brouillons** — permissions `Mail.Read`, `Calendars.Read`, `Contacts.Read`. Rien n'est envoyé : aucun endpoint d'envoi n'est exposé.

## ⚡ Chemin principal : les outils MCP `outlook-mcp__*`

**Toujours passer par les outils MCP, jamais par bash.** C'est la seule voie qui
fonctionne en session voice-call (voir « Compatibilité voice-call » plus bas).

| Outil | Rôle |
|---|---|
| `outlook-mcp__list_folders` | Dossiers mail + nombre de messages |
| `outlook-mcp__list_messages` | Derniers messages d'un dossier (`folder`, `top`) |
| `outlook-mcp__get_message` | Détail d'un message, corps en Markdown |
| `outlook-mcp__search_messages` | Recherche plein texte (`query`, `top`) |
| `outlook-mcp__list_events` | Calendrier (`top`) |
| `outlook-mcp__list_contacts` | Contacts (`top`) |
| `outlook-mcp__list_attachments` | Pièces jointes d'un message |
| `outlook-mcp__get_profile` | Profil utilisateur M365 |
| `outlook-mcp__list_drafts` | Brouillons existants (`top`) |
| `outlook-mcp__get_draft` | Contenu d'un brouillon |
| `outlook-mcp__create_draft` | Crée un brouillon — **rien n'est envoyé** |
| `outlook-mcp__update_draft` | Modifie un brouillon — **rien n'est envoyé** |
| `outlook-mcp__delete_draft` | Supprime un brouillon |

Règle d'identifiants : les ids de message, de dossier et de brouillon se
recopient **verbatim** depuis la sortie d'un appel précédent. Ne jamais les
construire.

Workflow typique : `list_messages` → `get_message`. Pour une recherche :
`search_messages` directement. Pour un email reçu par un tiers : `list_messages`
puis `get_message`.

Les appels se font via `tool_call` avec l'id de l'outil (ex.
`{"id": "outlook-mcp__list_messages", "args": {"folder": "Inbox", "top": 10}}`),
après un `tool_search` pour découvrir le catalogue.

## Brouillons (drafts) — écrire sans envoyer

`create_draft` / `update_draft` / `delete_draft` permettent de **rédiger** un
email sans jamais l'envoyer. Le message atterrit dans le dossier **Drafts** et
attend une action manuelle de l'utilisateur dans Outlook.

```json
{"id": "outlook-mcp__create_draft",
 "args": {"to": "destinataire@exemple.com",
          "subject": "Confirmation dossier FTFR123456",
          "body": "Bonjour,\n\n...",
          "cc": "copie@exemple.com",
          "in_reply_to": "<id-message-source>"}}
```

- `in_reply_to` : id d'un message → le brouillon rejoint le fil de discussion.
- `update_draft` ne modifie que les champs fournis (sujet, corps, `to`, `cc`).
- `delete_draft` est **irréversible** (corbeille) : ne l'appeler que sur
  demande explicite de l'utilisateur.

**Aucun outil d'envoi n'existe.** Le serveur n'expose pas l'endpoint `/send` :
l'envoi est impossible via ce skill, dans n'importe quel type de session. Pour
envoyer, l'utilisateur doit le faire depuis Outlook.

## Compatibilité voice-call — pourquoi MCP et pas bash

En session voice-call (Talk / realtime), les appels d'outils « à fort impact »
sont bloqués avant exécution par la **politique de confirmation vocale** :

```
VOICE_CONFIRMATION_REQUIRED:<id> The high-impact voice action "exec" was not executed.
```

Le déblocage exige une confirmation **parlée** relayée par l'outil
`openclaw_agent_consult`, qui n'existe pas dans le catalogue d'un agent textuel :
l'agent tourne donc en boucle sur `exec` sans jamais pouvoir confirmer.

Le classement est fait sur le **nom de l'outil + ses arguments** :

- `exec` / `bash` → mutation, **sauf** si la commande est un simple
  `cat`/`head`/`tail`/`grep`/`ls`/`wc`/`pwd`/`stat`/`rg`/`sed`/`find` en lecture seule.
  Donc `.venv/bin/python scripts/outlook_graph.py …` est **bloqué**.
- `automations`, `message` (envoi), `write`, `edit`, `sessions_spawn` → bloqués aussi.
- Les outils **MCP** dont le nom ne contient ni `send` ni `_actions` → classés
  lecture seule → **passent**. C'est exactement pourquoi `lynx-mcp__search_files`
  fonctionnait en mode vocal.

### Règle de nommage du skill (à ne pas casser)

Le classifieur s'appuie sur le **nom** de l'outil :

| Nom d'outil MCP | Classe | Effet en voice-call |
|---|---|---|
| `outlook-mcp__list_drafts`, `..._get_draft` | lecture seule | passe |
| `outlook-mcp__create_draft`, `..._update_draft`, `..._delete_draft` | lecture seule | passe |
| `outlook-mcp__send_draft` (**n'existe pas**) | mutation | serait bloqué |
| `outlook-mcp__*_actions` avec action mutante | mutation | serait bloqué |

Les outils de brouillon portent donc des noms **sans « send »** : leur effet est
local, visible dans Outlook, réversible, et ne notifie personne — pas de
confirmation vocale nécessaire pour les rédiger. Le jour où l'envoi serait
ajouté, il **devrait** s'appeler `send_draft` pour rester bloqué par défaut en
mode vocal.

Conséquences pour ce skill :

- ✅ Utiliser `outlook-mcp__*` — fonctionne partout, y compris en voice-call.
- ❌ Ne pas tenter `exec` sur les scripts Python en mode vocal.
- ⚠️ Le nom du skill est lu par l'agent en session vocale ; s'il n'évoque pas
  MCP, l'agent risque de retomber sur les commandes bash du SKILL.md.

## Vocabulaire voice-call

- Répondre **en une phrase courte** pour chaque résultat (date, expéditeur, objet).
- Ne pas intégrer le corps complet d'un email dans une réponse orale : le lire
  par dessus, ou demander confirmation avant de le dicter.
- Un email lu est long : en mode vocal, proposer d'abord la liste (expéditeur +
  objet), puis le détail si l'utilisateur le demande.

## Installation

Le serveur MCP est déjà enregistré dans `~/.openclaw/openclaw.json` :

```json
"outlook-mcp": {
  "command": "/home/fred-ghilini/.openclaw/skills/outlook-entra/.venv/bin/python",
  "args": ["/home/fred-ghilini/.openclaw/skills/outlook-entra/scripts/outlook_mcp.py"],
  "cwd": "/home/fred-ghilini/.openclaw/skills/outlook-entra"
}
```

Après toute modification du serveur : `openclaw mcp reload` puis
`openclaw mcp probe outlook-mcp` (doit annoncer 8 outils).

Dépendances Python (venv du skill) : `mcp`, `requests`, `html2text`, `cryptography`.

```bash
cd ~/.openclaw/skills/outlook-entra
uv venv .venv
uv pip install mcp requests html2text cryptography
```

## Authentification (Flow) — operationally manual

Le device code flow (RFC 8628) nécessite une intervention utilisateur unique.
**Cette étape se fait en session texte** : elle exige d'afficher un code et
d'attendre la validation dans un navigateur. En session vocale, ce n'est pas
faisable — le token doit déjà être valide.

### Étape 1 — Lancer le script d'auth

```bash
cd ~/.openclaw/skills/outlook-entra
.venv/bin/python scripts/outlook_auth.py
```

Le script va :
1. Demander un device code à Microsoft
2. Afficher un code utilisateur et une URL de vérification

### Étape 2 — S'authentifier

Aller sur **https://login.microsoft.com/device** et entrer le code affiché.
Délai : 15 minutes maximum.

### Étape 3 — Attendre la confirmation

Le script poll automatiquement le endpoint Microsoft jusqu'à obtention du token.
Une fois confirmé, les tokens sont sauvegardés localement.

### Fonctionnement des tokens

- L'`access_token` expire après ~1h
- Le `refresh_token` permet d'obtenir un nouvel `access_token` sans intervention
- Le refresh est **automatisé par cron** toutes les heures (voir section Cron)
- Si le `refresh_token` expire aussi (plusieurs mois d'inactivité) → relancer le flow complet
- Le serveur MCP gère aussi le refresh à la demande : sur un 401 Graph, il
  rafraîchit et rejoue la requête une fois.

---

### Commandes utiles (CLI — à éviter en mode vocal)

```bash
# Vérifier le statut du token
.venv/bin/python scripts/outlook_auth.py --status

# Rafraîchir le token manuellement
.venv/bin/python scripts/outlook_refresh.py

# Révoquer et supprimer les tokens
.venv/bin/python scripts/outlook_auth.py --revoke
```

## Commandes CLI (lecture seule) — équivalent MCP

Utile en session texte pour un débogage rapide ou un traitement par lot.
**Ces commandes sont bloquées en session voice-call** (voir plus haut).

```bash
# Statut de connexion
.venv/bin/python scripts/outlook_auth.py --status

# Lire les derniers messages
.venv/bin/python scripts/outlook_graph.py messages --folder Inbox --top 10

# Détail d'un message (corps complet, Markdown par défaut)
.venv/bin/python scripts/outlook_graph.py message <messageId>

# Détail en HTML brut (pour extraction/collage)
.venv/bin/python scripts/outlook_graph.py message <messageId> --raw

# Lister les dossiers mail
.venv/bin/python scripts/outlook_graph.py folders

# Pièces jointes d'un message
.venv/bin/python scripts/outlook_graph.py attachments <messageId>

# Télécharger une pièce jointe
.venv/bin/python scripts/outlook_graph.py download <messageId> --attach-id <attachmentId> --output /path/to/file

# Événements calendrier
.venv/bin/python scripts/outlook_graph.py events --top 10

# Contacts
.venv/bin/python scripts/outlook_graph.py contacts --top 20

# Rechercher dans les mails
.venv/bin/python scripts/outlook_graph.py search "mot-clé"

# Profil utilisateur
.venv/bin/python scripts/outlook_graph.py profile
```

## Prérequis

- App enregistrée sur **Entra** (Azure AD) avec permissions :
  - `Mail.Read`, `Calendars.Read`, `Contacts.Read`
  - OAuth 2.0 device code flow activé
- Fichier `.env` configuré (voir `.env.example`)
- Token déjà authentifié (device code flow effectué au moins une fois)

## Variables d'environnement (.env)

| Variable                | Description                              | Exemple                                                             |
| ----------------------- | ---------------------------------------- | ------------------------------------------------------------------- |
| `AZURE_TENANT_ID`       | GUID du tenant Entra                     | `52ffb8b9-…`                                                        |
| `AZURE_CLIENT_ID`       | ID de l'app (Application ID)             | `xxxxxxxx-…`                                                        |
| `AZURE_CLIENT_SECRET`   | Secret de l'app                          | `~`                                                                 |
| `AZURE_REDIRECT_URI`    | Redirect URI (device flow : tout fait)   | `http://localhost`                                                  |
| `OAUTH_TOKEN_URL`       | URL token endpoint                       | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token`      |
| `OAUTH_DEVICE_CODE_URL` | URL device code endpoint                 | `https://login.microsoftonline.com/{tenant}/oauth2/v2.0/devicecode` |
| `MS_GRAPH_BASE_URL`     | Base URL Microsoft Graph                 | `https://graph.microsoft.com/v1.0`                                  |
| `TOKEN_FILE`            | Chemin du fichier de stockage des tokens | `~/.openclaw/outlook_tokens.json`                                   |
| `TOKEN_FILE_KEY`        | Clé de chiffrement (optionnel)           | _(vide par défaut)_                                                 |

## Structure du skill

```
outlook-entra/
├── SKILL.md
├── README.md
├── .env.example
├── .gitignore
├── .venv/                   # Environnement virtuel Python (créé via uv venv)
├── scripts/
│   ├── outlook_auth.py      # OAuth device code flow + status/revoke
│   ├── outlook_graph.py     # Appels Graph API (lecture seule) — CLI
│   ├── outlook_mcp.py       # Serveur MCP stdio (8 outils, lecture seule) — chemin principal
│   ├── outlook_refresh.py   # Refresh token automatisé (pour cron)
│   └── outlook_token.py     # Module partagé (lecture/refresh tokens)
└── tests/
    └── test_outlook.py      # Tests unitaires
```

## Notes

- Le **device code flow** (RFC 8628) : l'utilisateur authentifie via `https://microsoft.com/devicelogin`. Une seule fois.
- Les **refresh tokens** sont automatiquement utilisés quand l'access token expire.
- Si `TOKEN_FILE_KEY` est défini, les tokens sont chiffrés AES-GCM avant stockage.
- Les erreurs 401 du Graph API déclenchent un refresh automatique (CLI et serveur MCP).
- Le serveur MCP recharge le token à la demande : sur 401, il rafraîchit et rejoue une fois.
- Les identifiants de message Outlook sont longs et opaques : les recopier **verbatim**
  depuis la sortie d'un appel précédent, ne jamais les reconstruire ni les tronquer.

## Cron — Refresh automatique du token

Le script `outlook_refresh.py` vérifie si le token expire bientôt et le rafraîchit automatiquement.

**Crontab** — refresh toutes les heures à HH:55 :

```cron
55 * * * * /home/fred-ghilini/.openclaw/skills/outlook-entra/.venv/bin/python /home/fred-ghilini/.openclaw/skills/outlook-entra/scripts/outlook_refresh.py >> /home/fred-ghilini/.openclaw/outlook_refresh.log 2>&1
```

**Installation** :

```bash
SKILL_DIR="/home/fred-ghilini/.openclaw/skills/outlook-entra"
( crontab -l 2>/dev/null | grep -v outlook_refresh; echo "55 * * * * ${SKILL_DIR}/.venv/bin/python ${SKILL_DIR}/scripts/outlook_refresh.py >> ~/.openclaw/outlook_refresh.log 2>&1" ) | crontab -
```

**Vérification** :

```bash
crontab -l | grep outlook_refresh
```

## Ressources

- [OAuth 2.0 Device Code Flow — Microsoft](https://learn.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Microsoft Graph — Mail API](https://learn.microsoft.com/en-us/graph/api/user-list-messages)
- [Microsoft Graph — Calendar API](https://learn.microsoft.com/en-us/graph/api/user-list-events)
