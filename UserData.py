import requests
import json
import openpyxl
from datetime import datetime

class UserData:
    def __init__(self, filename="users.json"):
        self.filename = filename

    def fetch_users_from_api(self, api_url):
        """Hakee käyttäjädatan rajapinnasta."""
        try:
            response = requests.get(api_url)
            response.raise_for_status()  # Raise an exception for HTTP errors
            return response.json()
        except requests.exceptions.RequestException as e:
            raise Exception(f"Virhe rajapintakutsussa: {e}")

    def save_users_to_file(self, users):
        """Tallentaa käyttäjädatan JSON-tiedostoon."""
        with open(self.filename, 'w', encoding='utf-8') as f:
            json.dump(users, f, indent=4, ensure_ascii=False)

    def load_users_from_file(self):
        """Lataa käyttäjädatan JSON-tiedostosta."""
        try:
            with open(self.filename, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            return None
        except json.JSONDecodeError:
            raise Exception(f"Virheellinen JSON-tiedosto: {self.filename}")

    def process_user_data(self, users):
        """Poimii ja muotoilee tarvittavat tiedot käyttäjistä."""
        processed_data = []
        for user in users:
            name_parts = user['name'].split()
            first_name = name_parts[0] if len(name_parts) > 0 else ""
            last_name = " ".join(name_parts[1:]) if len(name_parts) > 1 else ""
            processed_data.append({
                'sukunimi': last_name,
                'etunimi': first_name,
                'email': user['email'],
                'katuosoite': user['address']['street'],
                'postitoimipaikka': user['address']['city'],
                'postinumero': user['address']['zipcode'],
                'puhelin': user['phone'],
                'nettisivut': user['website']
            })
        return processed_data

    def sort_users(self, users):
        """Järjestää käyttäjät sukunimen ja etunimen mukaan."""
        return sorted(users, key=lambda x: (x['sukunimi'], x['etunimi']))

    def save_to_excel(self, users, filepath):
        """Tallentaa käyttäjädatan Excel-tiedostoon."""
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        header = ["sukunimi", "etunimi", "email", "katuosoite", "postitoimipaikka", "postinumero", "puhelin", "nettisivut"]
        sheet.append(header)
        for user in users:
            row = [user['sukunimi'], user['etunimi'], user['email'], user['katuosoite'],
                   user['postitoimipaikka'], user['postinumero'], user['puhelin'], user['nettisivut']]
            sheet.append(row)
        try:
            workbook.save(filepath)
        except Exception as e:
            raise Exception(f"Virhe Excel-tiedoston tallennuksessa: {e}")

def create_excel_filename():
    """Luo aikaleimalla varustetun Excel-tiedoston nimen."""
    now = datetime.now()
    return f"employees_{now.strftime('%Y%m%d%H%M%S')}.xlsx"