import json
import os
import random
import sys
import urllib.parse
import urllib.request
import urllib.error
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

    try:
        with urllib.request.urlopen(request) as response:
            return json.loads(response.read().decode("utf-8"))["access_token"]
    except urllib.error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(
            f"Entra token exchange failed with HTTP {exc.code}: {details}"
        ) from exc


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
    now = datetime.now(ZoneInfo("America/Chicago"))
    if not 9 <= now.hour < 11:
        raise RuntimeError(
            f"Teams posting blocked outside 9-11 AM Central: {now:%I:%M %p %Z}"
        )
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


def get_sent_event_keys(token, site_id):
    list_id = os.environ["CELEBRATION_LOG_LIST_ID"]
    url = (
        f"{GRAPH}/sites/{site_id}/lists/{list_id}/items"
        "?$expand=fields($select=EventKey)&$top=999"
    )
    items = get_all_items(token, url)
    return {
        item.get("fields", {}).get("EventKey")
        for item in items
        if item.get("fields", {}).get("EventKey")
    }


def record_sent_event(token, site_id, event, event_key, teams_sent, email_sent):
    list_id = os.environ["CELEBRATION_LOG_LIST_ID"]
    body = {
        "fields": {
            "Title": event["name"],
            "EventKey": event_key,
            "EmployeeEmail": event.get("email") or "",
            "EventType": (
                "Birthday"
                if event["type"] == "birthday"
                else "Anniversary"
            ),
            "EventDate": event["event_date"],
            "AnniversaryYears": event.get("years"),
            "TeamsSent": teams_sent,
            "EmailSent": email_sent,
            "ImageUrl": event["image"],
        }
    }

    request_json(
        f"{GRAPH}/sites/{site_id}/lists/{list_id}/items",
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        body=body,
    )

def send_anniversary_email(token, event):
    sender = os.environ["MAIL_SENDER"]
    recipient = event.get("email")

    if not recipient:
        raise RuntimeError(
            f"No employee email address is available for {event['name']}"
        )

    years = event["years"]
    first_name = event["name"].split()[0]

    subjects = {
        1: "Congratulations on Your One-Year Work Anniversary!",
        5: "Congratulations on Your Five-Year Work Anniversary!",
        10: "Congratulations on Your Ten-Year Work Anniversary!",
        15: "Celebrating Your 15 Year Work Anniversary!",
        20: "Celebrating Your Remarkable 20-Year Journey with Our Company",
    }

    messages = {
        1: (
            f"{first_name},<br><br>"
            "Happy 1st anniversary! It's been an incredible year with you "
            "as part of our team. Thank you for your hard work and dedication. "
            "Looking forward to many more milestones together!"
        ),
        5: (
            f"{first_name},<br><br>"
            "Today marks your fifth anniversary with Modern Business Interiors, "
            "and I wanted to express our appreciation for your hard work and "
            "dedication. Thank you for your continued commitment to excellence."
        ),
        10: (
            f"{first_name},<br><br>"
            "Happy 10th Anniversary! Your dedication, creativity, and hard work "
            "over the past decade have made a meaningful impact on our company. "
            "Thank you for being an integral part of our success."
        ),
        15: (
            f"{first_name},<br><br>"
            "Happy 15th Anniversary! Thank you for your outstanding dedication "
            "and contributions to Modern Business Interiors. Your commitment, "
            "expertise, and leadership have played an important role in our success."
        ),
        20: (
            f"{first_name},<br><br>"
            "Happy 20th Anniversary! Today we celebrate two decades of your "
            "incredible dedication and contributions to Modern Business Interiors. "
            "Thank you for your continued commitment and loyalty."
        ),
    }

    body = {
        "message": {
            "subject": subjects[years],
            "body": {
                "contentType": "HTML",
                "content": (
                    f"<p>{messages[years]}<br><br>"
                    "Best regards,<br><br>Jordan</p>"
                ),
            },
            "toRecipients": [
                {"emailAddress": {"address": recipient}}
            ],
            "bccRecipients": [
                {"emailAddress": {"address": sender}}
            ],
        },
        "saveToSentItems": True,
    }

    encoded_sender = urllib.parse.quote(sender)
    request_json(
        f"{GRAPH}/users/{encoded_sender}/sendMail",
        method="POST",
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        body=body,
    )

def anniversary_folder(years):
    milestones = {1, 5, 10, 15, 20, 25}
    return f"{years}-year" if years == 1 else (
        f"{years}-years" if years in milestones else "regular"
    )


def main():
    dry_run = os.getenv("DRY_RUN", "true").lower() == "true"
    now = datetime.now(ZoneInfo(os.getenv("TIME_ZONE", "America/Chicago")))

    github_event = os.getenv("GITHUB_EVENT_NAME")
    phase = "all"

    if github_event == "schedule":
        if 9 <= now.hour < 11:
            phase = "teams"
        elif now.hour >= 11:
            phase = "email"
        else:
            print(f"Before celebration window: {now:%I:%M %p %Z}")
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


    if os.getenv("PREVIEW_UPCOMING", "false").lower() == "true":
        from datetime import date, timedelta

        today = now.date()
        upcoming = []

        for item in items:
            fields = item.get("fields", {})
            if not active(fields.get("ActiveEmployee")):
                continue

            first = fields.get("PreferredFirstName") or fields.get("Title") or ""
            last = fields.get("field_1") or ""
            name = f"{first} {last}".strip()

            for kind, field in (
                ("Birthday", "field_2"),
                ("Anniversary", "field_3"),
            ):
                original = parse_date(fields.get(field))
                if not original:
                    continue

                for year in (today.year, today.year + 1, today.year + 2):
                    try:
                        occurrence = date(year, original.month, original.day)
                    except ValueError:
                        continue

                    if occurrence < today:
                        continue

                    if kind == "Anniversary" and year <= original.year:
                        continue

                    years = year - original.year
                    upcoming.append((occurrence, name, kind, years))
                    break

        upcoming.sort(key=lambda x: (x[0], x[1], x[2]))

        print(f"Active employees: {sum(active(i.get('fields', {}).get('ActiveEmployee')) for i in items)}")
        print("Upcoming celebrations:")

        for day, name, kind, years in upcoming[:10]:
            detail = f" ({years} years)" if kind == "Anniversary" else ""
            print(f"{day.isoformat()} | {kind} | {name}{detail}")

        if not upcoming:
            print("No upcoming celebrations found.")

        return

    events = []
    event_date = now.date().isoformat()

    for item in items:
        fields = item.get("fields", {})
        if not active(fields.get("ActiveEmployee")):
            continue

        first = fields.get("PreferredFirstName") or fields.get("Title") or ""
        last = fields.get("field_1") or ""
        name = f"{first} {last}".strip()
        email = fields.get("Email_x0020_Address") or ""

        birthday = parse_date(fields.get("field_2"))
        hire_date = parse_date(fields.get("field_3"))

        if birthday and (birthday.month, birthday.day) == (now.month, now.day):
            events.append(
                {
                    "type": "birthday",
                    "name": name,
                    "email": email,
                    "event_date": event_date,
                    "subject": f"Happy Birthday, {name}!",
                    "message": (
                        "Please take some time today to wish them "
                        "a happy birthday!"
                    ),
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
                        "email": email,
                        "send_email": years in {1, 5, 10, 15, 20},
                        "event_date": event_date,
                        "subject": f"Congratulations {name}, on {years} Years!",
                        "message": (
                            "Thank you for your exceptional dedication "
                            "and contributions!"
                        ),
                        "image": choose_image(Path("anniversary") / folder),
                    }
                )

    sent_keys = get_sent_event_keys(token, site["id"])

    print(f"Site: {site['displayName']}")
    print(
        "Active employee records read: "
        f"{sum(active(i.get('fields', {}).get('ActiveEmployee')) for i in items)}"
    )
    print(f"Celebrations found: {len(events)}")
    print(f"Phase: {phase}")
    print(f"Dry run: {dry_run}")

    for event in events:
        identity = event.get("email") or event["name"].lower()
        base_key = (
            f"{event['event_date']}|{event['type']}|{identity}"
        ).lower()
        teams_key = f"{base_key}|teams"
        email_key = f"{base_key}|email"

        print(
            f"{event['type']}: {event['name']} | "
            f"{event['image']}"
        )

        if phase in {"all", "teams"} and 9 <= datetime.now(ZoneInfo("America/Chicago")).hour < 11:
            if teams_key in sent_keys:
                print(f"Teams already sent: {event['name']}")
            elif not dry_run:
                post_teams(
                    event["subject"],
                    event["message"],
                    event["image"],
                )
                record_sent_event(
                    token,
                    site["id"],
                    event,
                    teams_key,
                    True,
                    False,
                )
                sent_keys.add(teams_key)
                print(f"Teams sent: {event['name']}")

        should_email = (
            event["type"] == "anniversary"
            and event.get("send_email", False)
        )

        if phase in {"all", "email"} and should_email:
            if email_key in sent_keys:
                print(f"Email already sent: {event['name']}")
            elif not dry_run:
                send_anniversary_email(token, event)
                record_sent_event(
                    token,
                    site["id"],
                    event,
                    email_key,
                    False,
                    True,
                )
                sent_keys.add(email_key)
                print(f"Jordan email sent: {event['name']}")

    if not events:
        print("No birthday or anniversary events found for today.")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise
