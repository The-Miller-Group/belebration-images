import json
import os
import random
import sys
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

GRAPH = "https://graph.microsoft.com/v1.0"
REPOSITORY = "The-Miller-Group/belebration-images"
BRANCH = "main"


def request_json(url, method="GET", headers=None, body=None):
    data = None if body is None else json.dumps(body).encode("utf-8")
    request = urllib.request.Request(
        url,
        data=data,
        method=method,
        headers=headers or {},
    )
    with urllib.request.urlopen(request) as response:
        content = response.read()
        return json.loads(content.decode("utf-8")) if content else {}


def get_graph_token():
    oidc_url = os.environ["ACTIONS_ID_TOKEN_REQUEST_URL"]
    oidc_token = os.environ["ACTIONS_ID_TOKEN_REQUEST_TOKEN"]

    separator = "&" if "?" in oidc_url else "?"
    token_response = request_json(
        f"{oidc_url}{separator}audience=api%3A%2F%2FAzureADTokenExchange",
        headers={"Authorization": f"Bearer {oidc_token}"},
    )

    assertion = token_response["value"]
    tenant_id = os.environ["AZURE_TENANT_ID"]
    client_id = os.environ["AZURE_CLIENT_ID"]

    form = urllib.parse.urlencode(
        {
            "client_id": client_id,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
            "client_assertion_type":
                "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
            "client_assertion": assertion,
        }
    ).encode("utf-8")

    request = urllib.request.Request(
        f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token",
        data=form,
        method="POST",
        headers={"Content-Type": "application/x-www-form-urlencoded"},
    )

    with urllib.request.urlopen(request) as response:
        return json.loads(response.read().decode("utf-8"))["access_token"]


def graph_get(token, path):
    return request_json(
        f"{GRAPH}{path}",
        headers={"Authorization": f"Bearer {token}"},
    )


def get_all_items(token, initial_url):
    items = []
    url = initial_url

    while url:
        result = request_json(
            url,
            headers={"Authorization": f"Bearer {token}"},
        )
        items.extend(result.get("value", []))
        url = result.get("@odata.nextLink")

    return items


def active(value):
    return value is True or str(value).lower() in {"1", "true", "yes"}


def parse_date(value):
    if not value:
        return None
    return datetime.fromisoformat(value.replace("Z", "+00:00"))


def image_url(relative_path):
    quoted = "/".join(
        urllib.parse.quote(part) for part in relative_path.parts
    )
    return (
        f"https://raw.githubusercontent.com/"
        f"{REPOSITORY}/{BRANCH}/{quoted}"
    )


def choose_image(folder):
    files = sorted(
        path for path in folder.iterdir()
        if path.is_file() and path.suffix.lower() in {".jpg", ".jpeg", ".png", ".gif"}
    )
    if not files:
        raise RuntimeError(f"No images found in {folder}")
    return image_url(random.choice(files))


def post_teams(subject, message, url):
    payload = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": None,
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.4",
                    "body": [
                        {
                            "type": "TextBlock",
                            "text": subject,
                            "weight": "Bolder",
                            "size": "Large",
                            "wrap": True,
                        },
                        {
                            "type": "Image",
                            "url": url,
                            "size": "Stretch",
                        },
                        {
                            "type": "TextBlock",
                            "text": message,
                            "weight": "Bolder",
                            "size": "Medium",
                            "wrap": True,
                            "spacing": "ExtraLarge",
                        },
                    ],
                },
            }
        ],
    }

    request_json(
        os.environ["TEAMS_WEBHOOK_URL"],
        method="POST",
        headers={"Content-Type": "application/json"},
        body=payload,
    )


def anniversary_folder(years):
    milestones = {1, 5, 10, 15, 20, 25}
    return f"{years}-year" if years == 1 else (
        f"{years}-years" if years in milestones else "regular"
    )


def main():
    dry_run = os.getenv("DRY_RUN", "true").lower() == "true"
    now = datetime.now(ZoneInfo(os.getenv("TIME_ZONE", "America/Chicago")))

    if os.getenv("GITHUB_EVENT_NAME") == "schedule" and now.hour != 9:
        print(f"Skipping scheduled run because local time is {now:%I:%M %p %Z}, not 9:00 AM.")
        return

    token = get_graph_token()

    host = os.environ["SHAREPOINT_HOST"]
    site_path = os.environ["SHAREPOINT_SITE_PATH"]
    site = graph_get(
        token,
        f"/sites/{host}:{site_path}?$select=id,displayName",
    )

    lists = graph_get(
        token,
        f"/sites/{site['id']}/lists?$select=id,displayName",
    )["value"]

    employee_list = next(
        item for item in lists
        if item["displayName"] == os.environ["EMPLOYEE_LIST_NAME"]
    )

    select = (
        "Title,field_1,PreferredFirstName,field_2,field_3,"
        "TerminationDate,ActiveEmployee,Email_x0020_Address"
    )
    url = (
        f"{GRAPH}/sites/{site['id']}/lists/{employee_list['id']}/items"
        f"?$expand=fields($select={select})&$top=999"
    )
    items = get_all_items(token, url)

    events = []

    for item in items:
        fields = item.get("fields", {})
        if not active(fields.get("ActiveEmployee")):
            continue

        first = fields.get("PreferredFirstName") or fields.get("Title") or ""
        last = fields.get("field_1") or ""
        name = f"{first} {last}".strip()

        birthday = parse_date(fields.get("field_2"))
        hire_date = parse_date(fields.get("field_3"))

        if birthday and (birthday.month, birthday.day) == (now.month, now.day):
            events.append(
                {
                    "type": "birthday",
                    "name": name,
                    "subject": f"Happy Birthday, {name}!",
                    "message": "Please take some time today to wish them a happy birthday!",
                    "image": choose_image(Path("birthday")),
                }
            )

        if hire_date and (hire_date.month, hire_date.day) == (now.month, now.day):
            years = now.year - hire_date.year
            if years > 0:
                folder = anniversary_folder(years)
                events.append(
                    {
                        "type": "anniversary",
                        "name": name,
                        "years": years,
                        "email": fields.get("Email_x0020_Address"),
                        "subject": f"Congratulations {name}, on {years} Years!",
                        "message": "Thank you for your exceptional dedication and contributions!",
                        "image": choose_image(Path("anniversary") / folder),
                    }
                )

    print(f"Site: {site['displayName']}")
    print(f"Active employee records read: {sum(active(i.get('fields', {}).get('ActiveEmployee')) for i in items)}")
    print(f"Celebrations found: {len(events)}")
    print(f"Dry run: {dry_run}")

    for event in events:
        print(
            f"{event['type']}: {event['name']} | "
            f"{event['image']}"
        )
        if not dry_run:
            post_teams(
                event["subject"],
                event["message"],
                event["image"],
            )

    if not events:
        print("No birthday or anniversary events found for today.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
