import re
import requests
import xmlrpc.client
import hashlib
import getpass
import sys

# Tapatalk Busker Organ
API_URL = (
    "https://www.tapatalk.com/groups/"
    "buskerorgan/mobiquo/mobiquo.php"
)


class Tapatalk:

    def __init__(self):
        self.session = requests.Session()
        self.session.headers.update({
            "Content-Type": "text/xml",
            "User-Agent": "Tapatalk Python Client"
        })

    
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
            timeout=30
        )

        response.raise_for_status()

        # XML ophalen
        xml = response.content.decode("utf-8")

        # Tapatalk gebruikt tijdzones die XML-RPC niet ondersteunt.
        # Verwijder de tijdzone-aanduiding uit dateTime velden.
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
            return result[0] if result else None

        except xmlrpc.client.Fault as e:
            raise RuntimeError(
                f"Tapatalk API fout: {e.faultString}"
            )

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

        print("\nVoorbereiden wachtwoord...")

        # Tapatalk kan MD5 of SHA1 voorschrijven
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
            raise RuntimeError(
                f"Onverwacht login antwoord: {result}"
            )

        if result.get("result") is not True:
            print("\nINLOGGEN MISLUKT")
            print("Reden:", result.get("result_text", "Onbekend"))
            return False

        print("\nINLOGGEN GELUKT!")

        if result.get("user_id"):
            print("User ID:", result["user_id"])

        print("Sessiecookies:",
              list(self.session.cookies.keys()))

        return True


def main():

    print("TAPATALK LOGIN TEST")
    print("===================")

    username = input("Gebruikersnaam: ")
    password = getpass.getpass("Wachtwoord: ")

    api = Tapatalk()

    try:

        config = api.get_config()

        if str(config.get("is_open", "1")) == "0":
            print("Tapatalk API is momenteel gesloten.")
            return

        if not api.login(username, password, config):
            return

        print("\nTest ophalen topic 775...")

        thread = api.call(
            "get_thread",
            "775",
            0
        )

        if isinstance(thread, dict):

            posts = thread.get("posts", [])

            print("Topic gevonden!")
            print("Aantal berichten:", len(posts))

            if posts:
                print("\nEerste bericht:")
                print(posts[0].get("post_title", ""))
                print(posts[0].get("post_content", ""))

        else:
            print("Onverwacht thread-antwoord:", thread)

    except Exception as e:
        print("\nFOUT:", e)


if __name__ == "__main__":
    main()
