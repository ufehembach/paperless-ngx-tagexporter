#!/usr/bin/env python3
import os
import requests
import pandas as pd
import argparse
from configparser import ConfigParser
from tqdm import tqdm

# Argumentparser einrichten
parser = argparse.ArgumentParser(
    description="Lade die Konfigurationsdatei für Paperless-ngx.")
parser.add_argument(
    "-c", "--config",
    default="my.ini",
    help="Pfad zur Konfigurationsdatei"
)
args = parser.parse_args()

# Konfigurationsdatei laden
config = ConfigParser()
config.read(args.config)

# Konfigurationseinstellungen auslesen
url = config['paperless']['url']
token = config['paperless']['token']
export_directory = config['paperless']['export_directory']
headers = {"Authorization": f"Token {token}"}

# Funktion zur Auflösung von IDs


def get_name_from_id(endpoint, id):
    response = requests.get(f"{url}/{endpoint}/{id}/", headers=headers)
    if response.status_code == 200:
        return response.json().get("name", "")
    return "Unbekannt"


# Custom-Field-Namen abrufen und ein Mapping erstellen
custom_fields_response = requests.get(
    f"{url}/custom_fields/", headers=headers)
custom_fields_map = {
    field["id"]: field["name"] for field in custom_fields_response.json()["results"]
}

# Tag ID ermitteln
tag_name = input("Bitte den Namen des Tags eingeben: ")
tags_response = requests.get(f"{url}/tags/", headers=headers)
tag_id = None
tag_dict = {tag["id"]: tag["name"] for tag in tags_response.json()["results"]}

for tag in tags_response.json()["results"]:
    if tag["name"].lower() == tag_name.lower():
        tag_id = tag["id"]
        break

if not tag_id:
    print("Tag nicht gefunden.")
    exit()

# Alle Dokumente abrufen und nach Tag filtern
documents = []
page = 1
while True:
    documents_response = requests.get(
        f"{url}/documents/?page_size=25&page={page}", headers=headers)
    if documents_response.status_code != 200:
        print(
            f"Fehler beim Abrufen der Dokumente. Status Code: {documents_response.status_code}")
        break

    data = documents_response.json()
    documents.extend([doc for doc in data["results"]
                     if tag_id in doc.get("tags", [])])
    if not data["next"]:
        break
    page += 1

# Daten für Excel sammeln
document_data = []
print("Dokumente werden verarbeitet...")

for doc in tqdm(documents, desc="Verarbeite Dokumente", unit="Dok", ncols=80):
    # Detailinformationen des Dokuments abrufen
    detailed_doc_response = requests.get(
        f"{url}/documents/{doc['id']}/", headers=headers)
    detailed_doc = detailed_doc_response.json()

    # Custom Fields auflösen
    custom_fields = {}
    if "custom_fields" in detailed_doc:
        for custom_field in detailed_doc["custom_fields"]:
            field_id = custom_field["field"]
            field_name = custom_fields_map.get(field_id, f"Feld {field_id}")
            custom_fields[field_name] = custom_field["value"]

    # Dokumentzeile für Excel-Liste vorbereiten
    row = {
        "ID": doc.get("id"),
        "Titel": doc.get("title"),
        "Datum": doc.get("created"),
        "Korrespondent": get_name_from_id("correspondents", doc.get("correspondent")),
        "Dokumenttyp": get_name_from_id("document_types", doc.get("document_type")),
        "Speicherpfad": get_name_from_id("storage_paths", doc.get("storage_path")),
        "Tags": ", ".join(tag_dict.get(tag_id, f"Tag {tag_id}") for tag_id in doc.get("tags", [])),
        **custom_fields
    }

    document_data.append(row)

# Exportverzeichnis erstellen
if not os.path.exists(export_directory):
    os.makedirs(export_directory)

# Excel-Datei erstellen
df = pd.DataFrame(document_data)
excel_path = os.path.join(export_directory, f"Dokumentenliste_{tag_name}.xlsx")
df.to_excel(excel_path, index=False)
print(f"\nExcel-Datei erfolgreich erstellt: {excel_path}")

print("Fertig!")
