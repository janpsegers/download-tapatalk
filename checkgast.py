import requests

API = "https://www.tapatalk.com/groups/buskerorgan/mobiquo/tapatalk.php"

r = requests.get(API, params={
    "method_name": "get_thread",
    "topicId": "775",
    "returnHtml": 0,
    "page": 1,
    "perPage": 5,
}, timeout=60)

print(r.status_code)
print(r.text[:500])