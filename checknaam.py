import requests, getpass, base64

API = "https://www.tapatalk.com/groups/buskerorgan/mobiquo/tapatalk.php"
naam = "janpsegers@kpnmail.nl"
ww = getpass.getpass("Wachtwoord: ")

b64 = lambda s: base64.b64encode(s.encode("utf-8")).decode("ascii")

varianten = {
    "GET gewoon":    ("get",  {"login_name": naam, "password": ww}),
    "GET base64":    ("get",  {"login_name": b64(naam), "password": b64(ww)}),
    "POST gewoon":   ("post", {"login_name": naam, "password": ww}),
    "POST base64":   ("post", {"login_name": b64(naam), "password": b64(ww)}),
    "POST username": ("post", {"username": naam, "password": ww}),
}

for titel, (methode, velden) in varianten.items():
    s = requests.Session()
    if methode == "get":
        r = s.get(API, params={"method_name": "login", **velden}, timeout=60)
    else:
        r = s.post(API, params={"method_name": "login"}, data=velden, timeout=60)
    print(titel, "->", r.status_code, r.text[:200])