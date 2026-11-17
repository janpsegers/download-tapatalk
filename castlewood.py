import re
import json
import time
import getpass
import hashlib
import xmlrpc.client
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from docx import Document
from docx.shared import Inches
from docx.image.image import Image as DocxImage

# --------------------------------------------------
# INSTELLINGEN
# --------------------------------------------------

BASE = "https://www.tapatalk.com/groups/buskerorgan/"
API_URL = urljoin(BASE, "mobiquo/mobiquo.php")

TOPIC_ID = "775"
PER_PAGE = 50

OUTPUT = Path("Castlewood_250")
IMAGE_DIR = OUTPUT / "afbeeldingen"
FILE_DIR = OUTPUT / "bijlagen"

OUTPUT.mkdir(exist_ok=True)
IMAGE_DIR.mkdir(exist_ok=True)
FILE_DIR.mkdir(exist_ok=True)


# --------------------------------------------------
# HULPFUNCTIES
# --------------------------------------------------

def decode(obj):
    """Zet bytes (base64-velden van Tapatalk) om naar tekst."""
    if isinstance(obj, bytes):
        return obj.decode("utf-8", errors="replace")
    if isinstance(obj, dict):
        return {k: decode(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [decode(v) for v in obj]
    return obj


def strip_control_chars(text):
    """Word accepteert geen stuurtekens in tekst."""
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)


def clean_text(content):
    """Haalt BBCode uit een bericht en laat leesbare tekst over."""
    if not content:
        return ""

    text = str(content)

    # Afbeeldingen komen apart in het document
    text = re.sub(r"\[img\].*?\[/img\]", "", text,
                  flags=re.I | re.S)

    # Citaten
    text = re.sub(r"\[quote[^\]]*\]", "\n« ", text, flags=re.I)
    text = re.sub(r"\[/quote\]", " »\n", text, flags=re.I)

    # Links: [url=X]tekst[/url] -> tekst (X)
    text = re.sub(
        r"\[url=([^\]]+)\](.*?)\[/url\]",
        r"\2 (\1)",
        text,
        flags=re.I | re.S
    )
    text = re.sub(r"\[url\](.*?)\[/url\]", r"\1", text,
                  flags=re.I | re.S)

    # Overige tags ([b], [i], [emoji328], [attachment=..] enz.)
    text = re.sub(r"\[/?[a-zA-Z*][^\]]*\]", "", text)

    return strip_control_chars(text).strip()


def add_picture_native_size(doc, path):
    """Zet een afbeelding in Word op zijn eigen formaat (96 pixels per inch).
    Is hij breder dan de pagina, dan wordt hij verkleind tot de paginabreedte."""
    section = doc.sections[0]
    max_width = section.page_width - section.left_margin - section.right_margin

    image = DocxImage.from_file(str(path))
    width = Inches(image.px_width / 96)

    if width > max_width:
        width = max_width

    doc.add_picture(str(path), width=width)


def get_image_urls(post):
    urls = []
    content = post.get("post_content", "") or ""

    # [img]-tags in het bericht
    for match in re.findall(r"\[img\](.*?)\[/img\]", content,
                            flags=re.I | re.S):
        urls.append(match.strip())

    # Bijlagen
    for key in ("attachments", "inlineattachments"):
        for attachment in post.get(key, []) or []:
            if not isinstance(attachment, dict):
                continue
            url = attachment.get("url") or attachment.get("thumbnail_url")
            if url:
                urls.append(url)

    result = []
    for url in urls:
        url = urljoin(BASE, url)
        if url not in result:
            result.append(url)

    return result


# --------------------------------------------------
# TAPATALK-CLIENT
# --------------------------------------------------

class Tapatalk:

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "text/xml",
            "User-Agent": "Tapatalk Python Client"
        })
        self.image_cache = {}
        self.image_number = 0
        self.file_number = 0

    def call(self, method, *args):

        request_xml = xmlrpc.client.dumps(
            args,
            methodname=method,
            allow_none=True,
            encoding="utf-8"
        )

        response = self.session.post(
            API_URL,
            data=request_xml.encode("utf-8"),
            timeout=60
        )

        response.raise_for_status()

        xml = response.content.decode("utf-8")

        # Tapatalk gebruikt tijdzones die XML-RPC niet ondersteunt.
        xml = re.sub(
            r'(<dateTime\.iso8601>[^<]*?)([+-]\d{2}:\d{2}|Z)(</dateTime\.iso8601>)',
            r'\1\3',
            xml
        )

        try:
            result, _ = xmlrpc.client.loads(
                xml,
                use_builtin_types=True
            )
            return decode(result[0]) if result else None

        except xmlrpc.client.Fault as e:
            raise RuntimeError(f"Tapatalk API fout: {e.faultString}")

        except Exception as e:
            raise RuntimeError(
                f"XML-parser fout: {e}\n"
                f"Antwoord: {xml[:500]}"
            )

    def get_config(self):
        print("\nOphalen API-configuratie...")

        config = self.call("get_config")

        if not isinstance(config, dict):
            raise RuntimeError(
                f"Onverwacht configuratieformaat: {config}"
            )

        print("API versie:", config.get("api_level"))
        print("Plugin versie:", config.get("version"))

        return config

    def login(self, username, password, config):

        if str(config.get("support_md5", "0")) == "1":
            password = hashlib.md5(
                password.encode("utf-8")
            ).hexdigest()
            print("Wachtwoord: MD5")

        elif str(config.get("support_sha1", "0")) == "1":
            password = hashlib.sha1(
                password.encode("utf-8")
            ).hexdigest()
            print("Wachtwoord: SHA1")

        else:
            print("Wachtwoord: ongewijzigd")

        print("\nInloggen...")

        result = self.call(
            "login",
            xmlrpc.client.Binary(username.encode("utf-8")),
            xmlrpc.client.Binary(password.encode("utf-8"))
        )

        if not isinstance(result, dict):
            raise RuntimeError(f"Onverwacht login antwoord: {result}")

        if result.get("result") is not True:
            print("\nINLOGGEN MISLUKT")
            print("Reden:", result.get("result_text", "Onbekend"))
            return False

        print("INLOGGEN GELUKT!")
        return True

    def get_thread_page(self, start):
        end = start + PER_PAGE - 1
        # get_thread(topic_id, start_num, last_num, return_html)
        return self.call("get_thread", TOPIC_ID, start, end, False)

    def download_image(self, url):
        # Dezelfde afbeelding niet twee keer downloaden
        if url in self.image_cache:
            return self.image_cache[url]

        try:
            response = self.session.get(
                url,
                headers={"Content-Type": None},
                timeout=40
            )
            response.raise_for_status()

            content_type = response.headers.get(
                "Content-Type", ""
            ).split(";")[0].lower()

            if not content_type.startswith("image/"):
                # Geen afbeelding (bijvoorbeeld pdf): als bijlage bewaren
                self.file_number += 1
                name = Path(urlparse(url).path).name or "bijlage"
                path = FILE_DIR / f"{self.file_number:03d}_{name}"
                path.write_bytes(response.content)
                print("  Bijlage opgeslagen:", path.name)
                self.image_cache[url] = ("bijlage", path)
                return ("bijlage", path)

            extensions = {
                "image/jpeg": ".jpg",
                "image/png": ".png",
                "image/gif": ".gif",
                "image/webp": ".webp"
            }
            extension = extensions.get(content_type, ".jpg")

            self.image_number += 1
            path = IMAGE_DIR / (
                f"afbeelding_{self.image_number:04d}{extension}"
            )
            path.write_bytes(response.content)

            self.image_cache[url] = ("afbeelding", path)
            return ("afbeelding", path)

        except Exception as error:
            print("  Afbeelding mislukt:", error)
            self.image_cache[url] = None
            return None


# --------------------------------------------------
# HOOFDPROGRAMMA
# --------------------------------------------------

def main():

    print("================================")
    print(" CASTLEWOOD 250 BUILD LOG")
    print(" Busker Organ Forum")
    print("================================")

    api = Tapatalk()
    config = api.get_config()

    if str(config.get("is_open", "1")) == "0":
        print("Tapatalk API is momenteel gesloten.")
        return

    # Inloggen is optioneel: het onderwerp is openbaar.
    print("\nLaat de gebruikersnaam leeg om als gast door te gaan.")
    username = input("Gebruikersnaam: ").strip()

    if username:
        password = getpass.getpass("Wachtwoord: ")
        if not api.login(username, password, config):
            antwoord = input("Doorgaan als gast? (j/n): ").strip().lower()
            if antwoord != "j":
                return
    else:
        print("Doorgaan als gast.")

    print("\nBerichten ophalen...")

    doc = Document()
    doc.add_heading("Castlewood 250 build log", 0)
    doc.add_paragraph("Busker Organ Forum")

    all_posts = []
    seen = set()
    total_posts = None
    start = 0

    while True:

        print(f"\nBerichten vanaf {start} ophalen...")

        thread = api.get_thread_page(start)

        if not isinstance(thread, dict):
            print("Onverwacht antwoord van forum:", thread)
            break

        if total_posts is None:
            total_posts = int(thread.get("total_post_num", 0))
            print("Totaal aantal berichten:", total_posts)
            print("Onderwerp:", thread.get("topic_title", ""))

        posts = thread.get("posts") or []

        if not posts:
            print("Geen berichten meer ontvangen.")
            break

        new_posts = 0

        for post in posts:

            post_id = str(post.get("post_id", ""))

            if post_id in seen:
                continue

            seen.add(post_id)
            new_posts += 1
            all_posts.append(post)

            number = len(all_posts)
            print(f"Bericht {number} van {total_posts}: {post_id}")

            doc.add_heading(f"Bericht {number}", level=2)

            title = post.get("post_title", "")
            if title:
                doc.add_paragraph("Titel: " + strip_control_chars(str(title)))

            doc.add_paragraph("Berichtnummer: " + post_id)
            doc.add_paragraph(
                "Auteur: " + strip_control_chars(
                    str(post.get("post_author_name", ""))
                )
            )
            doc.add_paragraph("Datum: " + str(post.get("post_time", "")))

            doc.add_paragraph(clean_text(post.get("post_content", "")))

            for image_url in get_image_urls(post):
                result = api.download_image(image_url)
                if not result:
                    continue
                soort, path = result
                if soort == "afbeelding":
                    try:
                        add_picture_native_size(doc, path)
                    except Exception:
                        print("  Niet in Word te zetten:", path.name)
                else:
                    doc.add_paragraph("Bijlage: " + path.name)

            doc.add_page_break()

        # Doorgaan vanaf het aantal dat we echt hebben ontvangen
        start += len(posts)

        if total_posts and start >= total_posts:
            break

        if new_posts == 0:
            print("Alleen al bekende berichten ontvangen, gestopt.")
            break

        time.sleep(1)

    # --------------------------------------------------
    # OPSLAAN
    # --------------------------------------------------

    print("\nResultaten opslaan...")

    with open(OUTPUT / "berichten.json", "w", encoding="utf-8") as file:
        json.dump(
            all_posts,
            file,
            ensure_ascii=False,
            indent=2,
            default=str
        )

    doc.save(OUTPUT / "Castlewood_250.docx")

    print("\n================================")
    print(" KLAAR")
    print("================================")
    print("Opgehaald:", len(all_posts), "berichten")
    print("Verwacht:", total_posts)
    print("Afbeeldingen:", api.image_number)
    print("Bijlagen:", api.file_number)
    print("Bestanden staan in:", OUTPUT.resolve())

    if total_posts and len(all_posts) != total_posts:
        print("\nLET OP: aantal berichten komt niet overeen!")


if __name__ == "__main__":

    try:
        main()

    except KeyboardInterrupt:
        print("\nAfgebroken door gebruiker.")

    except Exception as error:
        print("\nFOUT:")
        print(error)

    input("\nDruk op Enter om af te sluiten...")
