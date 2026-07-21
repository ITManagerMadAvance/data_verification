#!/usr/bin/env python3
"""
Vérification de données — Appel maintenance préventive
=========================================================

Applique les six dimensions du "Manuel de vérification de données" MadAvance
(Complétude, Promptitude, Validité, Unicité, Cohérence, Fiabilité) à
l'activité "Appel maintenance préventive", en croisant les données de
trois formulaires mWater :

  - Appel maintenance préventive
  - Maintenance préventive
  - Réparation après panne

... et d'un quatrième, utilisé uniquement pour la dimension Fiabilité :

  - Première réhabilitation (au sein du datagrid "Première réhabilitation",
    qui contient aussi d'autres types de travaux)

Le script télécharge les données via l'API mWater (datagrids déjà
configurés dans le portail), applique les règles, puis génère un rapport
Excel horodaté (data_verification_call_JJMMAAAA.xlsx) qu'il dépose sur
SharePoint via Microsoft Graph, et envoie un email de confirmation.

Inspiré du repo `ITManagerMadAvance/mWater_backup` pour les mécanismes
d'authentification mWater et Microsoft Graph.
"""

import csv
import io
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

import requests

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

MWATER_API_BASE = "https://api.mwater.co/v3"

# IDs des datagrids mWater (préconfigurés dans le portail, pas des secrets)
DATAGRID_APPEL = "5f443b8a8c144502a304c8c5c24d4f82"
DATAGRID_MAINTENANCE = "2cfe0ba7ac264d119cfc8964b5f3cebc"
DATAGRID_REPARATION = "d9f1c36a2d6340429658b6628fe81b88"
DATAGRID_REHAB = "0638d32971704164b9ac22d549d4818e"

SIGNAL_CODE_RE = re.compile(r"^([A-Z]+)_(\d{2})(\d{2})(\d{4})_([ES])(\d+)$")

DEPLOYMENT_PREFIX_MAP = {
    "MAR": ["Maroantsetra"],
    "FTU": ["Fort-Dauphin", "Taolagnaro"],
}


def raise_for_status_verbose(response):
    """Comme dans backup_mwater.py : loggue le corps complet en cas d'erreur HTTP."""
    if not response.ok:
        print(f"HTTP {response.status_code} sur {response.url}", file=sys.stderr)
        print(response.text[:2000], file=sys.stderr)
        response.raise_for_status()


# ---------------------------------------------------------------------------
# mWater : authentification et téléchargement des datagrids
# ---------------------------------------------------------------------------

def mwater_login(username, password):
    resp = requests.post(
        f"{MWATER_API_BASE}/clients",
        json={"username": username, "password": password},
        timeout=30,
    )
    raise_for_status_verbose(resp)
    client_id = resp.json().get("client")
    if not client_id:
        raise RuntimeError("Authentification mWater : champ 'client' absent de la réponse")
    return client_id


def download_datagrid(datagrid_id, client_id):
    """Télécharge un datagrid mWater et le retourne comme liste de dict (une par ligne)."""
    resp = requests.get(
        f"{MWATER_API_BASE}/datagrids/{datagrid_id}/download",
        params={"client": client_id, "share": "", "extraFilters": "[]", "format": "csv"},
        timeout=120,
    )
    raise_for_status_verbose(resp)
    # utf-8-sig pour gérer le BOM renvoyé par mWater
    text = resp.content.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))
    return list(reader)


# ---------------------------------------------------------------------------
# Utilitaires de parsing
# ---------------------------------------------------------------------------

def parse_dt(value):
    """Parse un horodatage mWater 'AAAA-MM-JJ HH:MM:SS'. Retourne None si vide/invalide."""
    value = (value or "").strip()
    if not value:
        return None
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


def parse_signal_code_date(code):
    """Extrait la date embarquée (JJMMAAAA) d'un signal code/reference. Retourne un datetime.date ou None."""
    m = SIGNAL_CODE_RE.match((code or "").strip())
    if not m:
        return None
    deployment, dd, mm, yyyy, _, _ = m.groups()
    try:
        return datetime(int(yyyy), int(mm), int(dd)).date()
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# Anomalie : structure commune
# ---------------------------------------------------------------------------

class Anomaly:
    def __init__(self, dimension, response_code, water_point_id, description, details=""):
        self.dimension = dimension
        self.response_code = response_code
        self.water_point_id = water_point_id
        self.description = description
        self.details = details

    def as_row(self):
        return [self.dimension, self.response_code, self.water_point_id, self.description, self.details]


# ---------------------------------------------------------------------------
# Dimension : Complétude
# ---------------------------------------------------------------------------

def check_completude(appel_rows):
    anomalies = []
    for r in appel_rows:
        status = r.get("Status", "")
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")
        if status == "Draft":
            anomalies.append(Anomaly(
                "Complétude", rc, wp,
                "Brouillon jamais soumis",
                f"Drafted On: {r.get('Drafted On', '')}",
            ))
        elif status == "Final" and not wp.strip():
            anomalies.append(Anomaly(
                "Complétude", rc, wp,
                "Réponse finalisée sans Water Point ID renseigné",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Promptitude
# ---------------------------------------------------------------------------

def check_promptitude(appel_rows):
    anomalies = []
    for r in appel_rows:
        drafted = parse_dt(r.get("Drafted On"))
        submitted = parse_dt(r.get("Submitted On"))
        if drafted and submitted and submitted < drafted:
            anomalies.append(Anomaly(
                "Promptitude", r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Submitted On antérieur à Drafted On",
                f"Drafted On: {drafted} / Submitted On: {submitted}",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Validité
# ---------------------------------------------------------------------------

def check_validite(appel_rows):
    anomalies = []
    for r in appel_rows:
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")
        signal_ref = r.get("Signal reference", "").strip()
        if signal_ref and not SIGNAL_CODE_RE.match(signal_ref):
            anomalies.append(Anomaly(
                "Validité", rc, wp,
                "Signal reference ne respecte pas le format attendu",
                f"Valeur : '{signal_ref}'",
            ))
        if wp and not wp.strip().isdigit():
            anomalies.append(Anomaly(
                "Validité", rc, wp,
                "Water Point ID non numérique",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Unicité
# ---------------------------------------------------------------------------

def check_unicite(appel_rows):
    anomalies = []
    counts = Counter(
        r.get("Signal reference", "").strip()
        for r in appel_rows
        if r.get("Signal reference", "").strip()
    )
    duplicated_codes = {code for code, n in counts.items() if n > 1}
    for r in appel_rows:
        signal_ref = r.get("Signal reference", "").strip()
        rejection = (r.get("Rejection message") or "").strip()
        if signal_ref in duplicated_codes:
            anomalies.append(Anomaly(
                "Unicité", r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Signal reference dupliqué",
                f"Code : {signal_ref} (x{counts[signal_ref]})",
            ))
        elif "doublon" in rejection.lower():
            anomalies.append(Anomaly(
                "Unicité", r.get("Response Code", ""), r.get("Water Point ID > Unique ID", ""),
                "Rejet mWater pour doublon de signal code",
                f"Message : {rejection}",
            ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Cohérence
# ---------------------------------------------------------------------------

def index_by_signal_code(rows, field="Signal code"):
    index = defaultdict(list)
    for r in rows:
        code = (r.get(field) or "").strip()
        if code:
            index[code].append(r)
    return index


def check_coherence(appel_rows, maintenance_rows, reparation_rows):
    anomalies = []
    maintenance_idx = index_by_signal_code(maintenance_rows)
    reparation_idx = index_by_signal_code(reparation_rows)

    for r in appel_rows:
        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "")
        signal_ref = r.get("Signal reference", "").strip()
        deployment = (r.get("Deployment") or "").strip()
        pump_state = (r.get("Is the pump currently working ?") or "").strip()

        if not signal_ref:
            continue

        m = SIGNAL_CODE_RE.match(signal_ref)
        if not m:
            continue  # déjà signalé en Validité
        code_prefix = m.group(1)

        # 1. Préfixe de déploiement
        if deployment and code_prefix != deployment:
            anomalies.append(Anomaly(
                "Cohérence", rc, wp,
                "Préfixe du signal code différent du déploiement déclaré",
                f"Signal reference : {signal_ref} / Deployment : {deployment}",
            ))

        # 2. Présence dans le bon formulaire selon l'état de la pompe
        found_in_maintenance = signal_ref in maintenance_idx
        found_in_reparation = signal_ref in reparation_idx

        if pump_state == "Partially" and not found_in_maintenance:
            anomalies.append(Anomaly(
                "Cohérence", rc, wp,
                "Pompe 'Partiellement' fonctionnelle mais signal code absent du formulaire Maintenance préventive",
                f"Signal reference : {signal_ref}",
            ))
        elif pump_state == "No" and not found_in_reparation:
            anomalies.append(Anomaly(
                "Cohérence", rc, wp,
                "Pompe 'Non' fonctionnelle mais signal code absent du formulaire Réparation après panne",
                f"Signal reference : {signal_ref}",
            ))

        # 3. Date du signal code <= date de l'activité (Completion date of the work)
        code_date = parse_signal_code_date(signal_ref)
        if code_date is None:
            continue

        matches = maintenance_idx.get(signal_ref, []) + reparation_idx.get(signal_ref, [])
        for match in matches:
            completion = parse_dt(match.get("Completion date of the work"))
            if completion and code_date > completion.date():
                anomalies.append(Anomaly(
                    "Cohérence", rc, wp,
                    "Date du signal code postérieure à la date de l'activité",
                    f"Signal code : {signal_ref} (date : {code_date}) / "
                    f"Completion date of the work : {completion.date()}",
                ))
    return anomalies


# ---------------------------------------------------------------------------
# Dimension : Fiabilité
# ---------------------------------------------------------------------------

def check_fiabilite(appel_rows, rehab_rows):
    anomalies = []

    wp_field = "De quel point d'eau s'agit-il? > Unique ID"

    # Sous-ensemble "Première réhabilitation" (le datagrid contient d'autres types de travaux)
    rehab_only = [r for r in rehab_rows if (r.get("Type de travaux") or "").strip() == "Première réhabilitation"]

    # Source de vérité : réhabilitations réussies (Final, pas Rejected)
    success_wp_ids = {
        r.get(wp_field, "").strip()
        for r in rehab_only
        if r.get("Status") == "Final" and r.get(wp_field, "").strip()
    }

    # Contexte complet (tous types de travaux, tous statuts) pour diagnostic
    all_wp_context = defaultdict(list)
    for r in rehab_rows:
        wp = r.get(wp_field, "").strip()
        if wp:
            all_wp_context[wp].append((r.get("Type de travaux", ""), r.get("Status", "")))

    for r in appel_rows:
        rejection = (r.get("Rejection message") or "").strip()
        if "id incorrect" not in rejection.lower():
            continue

        rc = r.get("Response Code", "")
        wp = r.get("Water Point ID > Unique ID", "").strip()

        if wp in success_wp_ids:
            anomalies.append(Anomaly(
                "Fiabilité", rc, wp,
                "Rejet 'ID incorrect' potentiellement injustifié",
                f"Ce Water Point ID existe dans la liste des Premières réhabilitations réussies (Status Final)",
            ))
        elif wp in all_wp_context:
            context = "; ".join(f"{t or 'type inconnu'} ({s})" for t, s in all_wp_context[wp])
            anomalies.append(Anomaly(
                "Fiabilité", rc, wp,
                "Rejet 'ID incorrect' à réexaminer — WP ID trouvé ailleurs dans le datagrid Réhabilitation",
                f"Trouvé sous : {context}",
            ))
        # Sinon : WP ID introuvable dans le datagrid réhabilitation -> rejet probablement justifié, pas d'anomalie.

    return anomalies


# ---------------------------------------------------------------------------
# Rapport Excel
# ---------------------------------------------------------------------------

def build_report(anomalies, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.title = "Anomalies"
    headers = ["Dimension", "Response Code", "Water Point ID", "Description", "Détails"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)

    for a in anomalies:
        ws.append(a.as_row())

    # Onglet résumé par dimension
    summary = wb.create_sheet("Résumé")
    summary.append(["Dimension", "Nombre d'anomalies"])
    for cell in summary[1]:
        cell.font = Font(bold=True)
    counts = Counter(a.dimension for a in anomalies)
    for dimension in ["Complétude", "Promptitude", "Validité", "Unicité", "Cohérence", "Fiabilité"]:
        summary.append([dimension, counts.get(dimension, 0)])
    summary.append(["Total", len(anomalies)])

    wb.save(output_path)


# ---------------------------------------------------------------------------
# Microsoft Graph : upload SharePoint + email
# ---------------------------------------------------------------------------

def graph_token(tenant_id, client_id, client_secret):
    resp = requests.post(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=30,
    )
    raise_for_status_verbose(resp)
    return resp.json()["access_token"]


def upload_to_sharepoint(token, drive_id, folder_item_id, file_path, file_name):
    with open(file_path, "rb") as f:
        content = f.read()
    resp = requests.put(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_item_id}:/{file_name}:/content",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        },
        data=content,
        timeout=120,
    )
    raise_for_status_verbose(resp)
    return resp.json()


def send_confirmation_email(token, sender, recipients, file_name, anomaly_count, counts_by_dimension):
    lines = "".join(f"<li>{dim} : {n}</li>" for dim, n in counts_by_dimension.items())
    body_html = (
        f"<p>Vérification de données \"Appel maintenance préventive\" exécutée.</p>"
        f"<p>Total anomalies détectées : <b>{anomaly_count}</b></p>"
        f"<ul>{lines}</ul>"
        f"<p>Rapport déposé sur SharePoint : <b>{file_name}</b></p>"
    )
    resp = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail",
        headers={"Authorization": f"Bearer {token}"},
        json={
            "message": {
                "subject": f"Vérification de données — {file_name}",
                "body": {"contentType": "HTML", "content": body_html},
                "toRecipients": [{"emailAddress": {"address": addr.strip()}} for addr in recipients.split(",")],
            }
        },
        timeout=30,
    )
    raise_for_status_verbose(resp)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    mwater_username = os.environ["MWATER_USERNAME"]
    mwater_password = os.environ["MWATER_PASSWORD"]

    azure_tenant_id = os.environ["AZURE_TENANT_ID"]
    azure_client_id = os.environ["AZURE_CLIENT_ID"]
    azure_client_secret = os.environ["AZURE_CLIENT_SECRET"]

    sharepoint_drive_id = os.environ["SHAREPOINT_DRIVE_ID"]
    sharepoint_folder_item_id = os.environ["SHAREPOINT_FOLDER_ITEM_ID"]

    email_sender = os.environ["EMAIL_SENDER"]
    email_recipients = os.environ["EMAIL_RECIPIENTS"]

    print("Authentification mWater...")
    client_id = mwater_login(mwater_username, mwater_password)

    print("Téléchargement des datagrids...")
    appel_rows = download_datagrid(DATAGRID_APPEL, client_id)
    maintenance_rows = download_datagrid(DATAGRID_MAINTENANCE, client_id)
    reparation_rows = download_datagrid(DATAGRID_REPARATION, client_id)
    rehab_rows = download_datagrid(DATAGRID_REHAB, client_id)
    print(f"  Appel: {len(appel_rows)} / Maintenance: {len(maintenance_rows)} / "
          f"Réparation: {len(reparation_rows)} / Réhabilitation: {len(rehab_rows)}")

    print("Application des règles de vérification...")
    anomalies = []
    anomalies += check_completude(appel_rows)
    anomalies += check_promptitude(appel_rows)
    anomalies += check_validite(appel_rows)
    anomalies += check_unicite(appel_rows)
    anomalies += check_coherence(appel_rows, maintenance_rows, reparation_rows)
    anomalies += check_fiabilite(appel_rows, rehab_rows)
    print(f"  {len(anomalies)} anomalies détectées")

    file_name = f"data_verification_call_{datetime.now().strftime('%d%m%Y')}.xlsx"
    output_path = f"/tmp/{file_name}"
    build_report(anomalies, output_path)
    print(f"Rapport généré : {output_path}")

    print("Authentification Microsoft Graph...")
    token = graph_token(azure_tenant_id, azure_client_id, azure_client_secret)

    print("Dépôt du rapport sur SharePoint...")
    upload_to_sharepoint(token, sharepoint_drive_id, sharepoint_folder_item_id, output_path, file_name)

    print("Envoi de l'email de confirmation...")
    counts_by_dimension = Counter(a.dimension for a in anomalies)
    send_confirmation_email(token, email_sender, email_recipients, file_name, len(anomalies), counts_by_dimension)

    print("Terminé.")


if __name__ == "__main__":
    main()
