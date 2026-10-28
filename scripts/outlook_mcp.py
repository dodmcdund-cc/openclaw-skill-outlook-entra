#!/usr/bin/env python3
"""
outlook_mcp.py — Serveur MCP (stdio) pour Outlook / Microsoft Graph (lecture seule).

Pourquoi un serveur MCP plutôt que des scripts bash ?
-------------------------------------------------------------
En session voice-call (Talk / realtime), les appels d'outils « à fort impact »
sont bloqués par la politique de confirmation vocale (VOICE_CONFIRMATION_REQUIRED).
`exec` est classé comme mutation dès que la commande n'est pas un simple
`cat/head/tail/grep/ls` en lecture seule — donc les scripts Python du skill
échouent en mode vocal, alors que les outils MCP (nom `outlook-mcp__*`) sont
classés lecture seule et passent.

Ce serveur expose les mêmes opérations en lecture seule, via des appels d'outils
MCP. Il n'expose AUCUNE écriture : Mail.Read, Calendars.Read, Contacts.Read.

Outils :
  list_folders        — dossiers mail + nombre de messages
  list_messages       — derniers messages d'un dossier
  get_message         — détail d'un message (corps en Markdown)
  search_messages     — recherche dans les mails
  list_events         — événements du calendrier
  list_contacts       — contacts
  list_attachments    — pièces jointes d'un message
  get_profile         — profil utilisateur

Usage (stdio) :
  .venv/bin/python scripts/outlook_mcp.py
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

import html2text
import requests
from mcp.server.mcpserver import MCPServer
from mcp.types import ToolAnnotations

import outlook_token as ot

GRAPH_BASE = ot.GRAPH_BASE
TIMEOUT = ot.REQUEST_TIMEOUT


# ── HTTP Graph ────────────────────────────────────────────────────────────────

def _graph(path: str, params: dict | None = None, token: str | None = None, retry: bool = True):
    """GET Graph avec refresh automatique sur 401 (token expiré/révoqué)."""
    token = token or ot.ensure_valid_token()
    url = f"{GRAPH_BASE}{path}"
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}

    resp = requests.get(url, headers=headers, params=params, timeout=TIMEOUT)

    if resp.status_code == 401 and retry:
        # Token révoqué ou expiré : on force un refresh puis on rejoue une fois.
        try:
            ot.refresh_token()
        except Exception:
            pass
        return _graph(path, params=params, retry=False)

    if resp.status_code >= 400:
        try:
            err = resp.json()
            msg = err.get("error", {}).get("message", err.get("error", "Unknown"))
            code = err.get("error", {}).get("code", "")
        except Exception:
            msg = resp.text[:200]
            code = ""
        if resp.status_code == 401:
            raise RuntimeError(
                "Graph API 401 — token invalidé. "
                "Relancez l'authentification : python scripts/outlook_auth.py"
            )
        raise RuntimeError(f"Graph API {resp.status_code} ({code}): {msg}")

    return resp.json()


# ── Formatters (sortie compacte, lisible à l'oral) ───────────────────────────

def _sender(msg: dict) -> str:
    return msg.get("from", {}).get("emailAddress", {}).get("name") or \
           msg.get("from", {}).get("emailAddress", {}).get("address") or "?"


def _fmt_message_line(m: dict) -> str:
    flags = "".join([
        "" if m.get("isRead") else "•",
        "📎" if m.get("hasAttachments") else "",
    ])
    when = (m.get("receivedDateTime") or "")[:16].replace("T", " ")
    return (
        f"{flags} {when} | {_sender(m)} | {m.get('subject') or '(sans objet)'}\n"
        f"   id={m.get('id')}"
    )


def _html_to_md(raw_html: str) -> str:
    h2t = html2text.HTML2Text()
    h2t.body_width = 0
    h2t.ignore_links = False
    return h2t.handle(raw_html).strip()


# ── Serveur MCP ───────────────────────────────────────────────────────────────

# Tous les outils sont en lecture seule : le hint explicite évite que les
# runtimes MCP exigeant des annotations déclenchent une approbation par appel.
READ_ONLY = ToolAnnotations(read_only_hint=True, destructive_hint=False,
                            idempotent_hint=True, open_world_hint=True)

server = MCPServer(
    name="outlook-mcp",
    version="1.0",
    instructions=(
        "Outlook / Microsoft Graph — LECTURE SEULE (Mail.Read, Calendars.Read, Contacts.Read). "
        "Workflow : list_folders ou list_messages pour repérer un message, get_message pour le "
        "corps complet, search_messages pour une recherche plein texte, list_events pour le "
        "calendrier, list_attachments pour les pièces jointes. "
        "Les identifiants de message et de dossier se recopient VERBATIM depuis la sortie d'un "
        "appel précédent ; ne jamais les construire. "
        "Aucun outil n'écrit : l'envoi, la suppression et le marquage lu sont hors périmètre."
    ),
)


@server.tool(
    name="list_folders",
    description="Liste les dossiers mail Outlook avec le nombre de messages. Retourne displayName, id, totalItemCount.",
    annotations=READ_ONLY,
)
def list_folders() -> str:
    data = _graph("/me/mailFolders", params={"$select": "id,displayName,totalItemCount"})
    lines = [
        f"- {f.get('displayName')} (id={f.get('id')}) — {f.get('totalItemCount', '?')} messages"
        for f in data.get("value", [])
    ]
    return "\n".join(lines) or "(aucun dossier)"


@server.tool(
    name="list_messages",
    description=(
        "Derniers messages d'un dossier Outlook. "
        "folder: nom de dossier (Inbox par défaut) ou id de dossier recopié verbatim. "
        "top: nombre de messages (défaut 10, max 50). "
        "Retourne date, expéditeur, objet et id de message."
    ),
    annotations=READ_ONLY,
)
def list_messages(folder: str = "Inbox", top: int = 10) -> str:
    top = max(1, min(int(top or 10), 50))
    data = _graph(
        f"/me/mailFolders/{folder}/messages",
        params={
            "$top": top,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments",
        },
    )
    msgs = data.get("value", [])
    if not msgs:
        return f"(aucun message dans « {folder} »)"
    return "\n".join(_fmt_message_line(m) for m in msgs)


@server.tool(
    name="get_message",
    description=(
        "Détail complet d'un message Outlook : expéditeur, destinataires, date, corps "
        "(converti en Markdown) et pièces jointes. "
        "message_id: id recopié verbatim depuis list_messages ou search_messages."
    ),
    annotations=READ_ONLY,
)
def get_message(message_id: str) -> str:
    m = _graph(
        f"/me/messages/{message_id}",
        params={
            "$select": "id,subject,from,toRecipients,receivedDateTime,isRead,body,hasAttachments,webLink"
        },
    )
    to = ", ".join(
        r.get("emailAddress", {}).get("address", "?") for r in m.get("toRecipients", [])
    )
    head = [
        f"De    : {m.get('from', {}).get('emailAddress', {}).get('address', '?')}",
        f"À     : {to}",
        f"Date  : {m.get('receivedDateTime', '?')}",
        f"Lu    : {'oui' if m.get('isRead') else 'non'}",
        f"Objet : {m.get('subject') or '(sans objet)'}",
        f"id    : {m.get('id')}",
    ]
    body = m.get("body", {})
    raw = body.get("content") or "(vide)"
    if "html" in (body.get("contentType") or "").lower():
        raw = _html_to_md(raw) or raw
    return "\n".join(head) + "\n\n--- Corps ---\n" + raw


@server.tool(
    name="search_messages",
    description=(
        "Recherche plein texte dans les mails Outlook (sujet et corps). "
        "query: terme recherché, entre guillemets. top: nombre de résultats (défaut 10, max 50)."
    ),
    annotations=READ_ONLY,
)
def search_messages(query: str, top: int = 10) -> str:
    top = max(1, min(int(top or 10), 50))
    data = _graph(
        "/me/messages",
        params={
            "$top": top,
            "$search": f'"{query}"',
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments",
        },
    )
    msgs = data.get("value", [])
    if not msgs:
        return f"(aucun résultat pour « {query} »)"
    return "\n".join(_fmt_message_line(m) for m in msgs)


@server.tool(
    name="list_events",
    description=(
        "Événements du calendrier Outlook, triés par date de début. "
        "top: nombre d'événements (défaut 10, max 50). "
        "Retourne sujet, début, fin, lieu, nombre de participants et id."
    ),
    annotations=READ_ONLY,
)
def list_events(top: int = 10) -> str:
    top = max(1, min(int(top or 10), 50))
    data = _graph(
        "/me/calendar/events",
        params={
            "$top": top,
            "$orderby": "start/dateTime asc",
            "$select": "id,subject,start,end,location,attendees,isAllDay",
        },
    )
    events = data.get("value", [])
    if not events:
        return "(aucun événement)"
    out = []
    for e in events:
        start = (e.get("start", {}).get("dateTime") or "?")[:16].replace("T", " ")
        end = (e.get("end", {}).get("dateTime") or "?")[:16].replace("T", " ")
        out.append(
            f"📅 {start} → {end} | {e.get('subject') or '(sans objet)'}\n"
            f"   lieu={e.get('location', {}).get('displayName') or '-'} "
            f"| participants={len(e.get('attendees', []))}\n"
            f"   id={e.get('id')}"
        )
    return "\n".join(out)


@server.tool(
    name="list_contacts",
    description="Contacts Outlook (nom, prénom, adresse e-mail). top: nombre de contacts (défaut 20, max 100).",
    annotations=READ_ONLY,
)
def list_contacts(top: int = 20) -> str:
    top = max(1, min(int(top or 20), 100))
    data = _graph(
        "/me/contacts",
        params={
            "$top": top,
            "$orderby": "givenName",
            "$select": "id,givenName,surname,emailAddresses,companyName",
        },
    )
    contacts = data.get("value", [])
    if not contacts:
        return "(aucun contact)"
    out = []
    for c in contacts:
        name = f"{c.get('givenName', '')} {c.get('surname', '')}".strip() or "?"
        email = (c.get("emailAddresses") or [{}])[0].get("address", "?")
        company = c.get("companyName") or ""
        out.append(f"- {name} <{email}>{' — ' + company if company else ''}")
    return "\n".join(out)


@server.tool(
    name="list_attachments",
    description=(
        "Liste les pièces jointes d'un message Outlook (nom, taille, type). "
        "message_id: id recopié verbatim depuis list_messages. "
        "La PII ne doit jamais être journalisée."
    ),
    annotations=READ_ONLY,
)
def list_attachments(message_id: str) -> str:
    data = _graph(
        f"/me/messages/{message_id}/attachments",
        params={"$select": "id,name,size,contentType"},
    )
    items = data.get("value", [])
    if not items:
        return "(aucune pièce jointe)"
    return "\n".join(
        f"- {a.get('name')} ({round((a.get('size') or 0) / 1024, 1)} Ko, {a.get('contentType')})\n"
        f"  id={a.get('id')}"
        for a in items
    )


@server.tool(
    name="get_profile",
    description="Profil de l'utilisateur Microsoft 365 connecté : nom, adresse e-mail, id Entra.",
    annotations=READ_ONLY,
)
def get_profile() -> str:
    me = _graph("/me")
    return (
        f"Nom      : {me.get('displayName', '?')}\n"
        f"Email    : {me.get('mail') or me.get('userPrincipalName', '?')}\n"
        f"ID Entra : {me.get('id', '?')}"
    )


def main() -> None:
    # stdout = protocole MCP uniquement ; les logs vont sur stderr.
    print("[outlook-mcp] ready — transport=stdio, lecture seule", file=sys.stderr)
    server.run(transport="stdio")


if __name__ == "__main__":
    main()
