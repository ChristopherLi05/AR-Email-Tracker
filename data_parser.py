import datetime
import json
import random
import re
import os
import csv

import bs4.element
from libratom.lib.pff import PffArchive
import pypff
from bs4 import BeautifulSoup


class Person:
    def __init__(self, name: str, preferred_name: str, emails: list[str]):
        self.name = (name or "").strip()
        self.preferred_name = (preferred_name or "").strip()

        self.emails = {i.strip() for i in emails}
        self.sanitized_emails = {i.lower().replace(".", "") for i in self.emails}

    def __repr__(self):
        return f"{self.__class__.__name__}(name={self.name}, preferred_name={self.preferred_name}, emails={self.emails})"

    def _get_sanitized_names(self):
        return re.sub(r"[^a-z\s]", "", self.name.lower()).split(), re.sub(r"[^a-z\s]", "",
                                                                          self.preferred_name.lower()).split()

    def add_email(self, email):
        self.emails.add(email)
        self.sanitized_emails.add(email.lower().replace(".", ""))

    def is_same_person(self, other: 'Person'):
        if self.does_email_match(other):
            return True

        return self.does_name_match(other)

    def does_email_match(self, other: 'Person'):
        return not self.sanitized_emails.isdisjoint(other.sanitized_emails)

    def does_name_match(self, other: 'Person'):
        # Basic Name Matching ; We know that there's at least 1 intersection if the length < 4
        # # if not {self.name, self.preferred_name}.isdisjoint({other.name, other.preferred_name}):
        # #     return True
        if self.name == other.name or self.preferred_name == other.preferred_name:
            return True

        # for name1, name2 in zip(self._get_sanitized_names(), other._get_sanitized_names()):
        #     if name1 == name2 and name1:
        #         return True
        #     # First + Last Name matching
        #     # if name1[0] == name2[0] and name1[-1] == name2[-1]:
        #     #     return True

        return False


class EmailMessage(Person):
    def __init__(self, message: pypff.message):
        if match := re.search(r"From: (.+? )?<(.+?)>", message.transport_headers or ""):
            sender_email = [match.group(2)]
        elif match := re.search(r"From: ([^ ]+?@[^ ]+?\.[^ ]+?)", message.transport_headers or ""):
            sender_email = [match.group(1)]
        else:
            sender_email = []
        super().__init__(message.sender_name, message.sender_name, sender_email)

        self.receive_time = int(message.get_delivery_time().timestamp() * 1000)
        self.email_contents = EmailMessage.parse_html(message.html_body)

    def __repr__(self):
        return f"{self.__class__.__name__}(name={self.name}, preferred_name={self.preferred_name}, emails={self.emails}, receive_time={self.receive_time})"

    def does_name_match(self, other: 'Person'):
        # TODO - figure out if this is good or not
        if self.name == other.preferred_name:
            return True

        return Person.does_name_match(self, other)

    @staticmethod
    def parse_html(html):
        parsed_html = BeautifulSoup((html or b"").decode("latin-1").strip().split("<hr")[0].split("---- Replied Message ----")[0], "html.parser")

        # Jank way of removing gmail quotes
        to_remove = [
            ("div", {"class": "gmail_quote"}),
            ("div", {"id": "mail-editor-reference-message-container"}),
            ("blockquote", {"id": "isReplyContent"}),
            ("blockquote", {"type": "cite"}),
        ]

        for tag, atr in to_remove:
            for elem in parsed_html.find_all(tag, atr):
                elem.decompose()

        return re.sub("\n\\s+", "\n", parsed_html.get_text()).strip()


def extract_emails(email_export_file: str):
    if not os.path.exists(email_export_file):
        print("Could not find file")
        return []
    elif not email_export_file.endswith(".pst"):
        print("Inputted file is not a .pst file")
        return []

    messages = []
    archive = PffArchive(os.path.abspath(email_export_file))

    for folder in archive.folders():
        if folder.name in [None]:
            continue

        for m in folder.sub_messages:
            messages.append(EmailMessage(m))

    return messages


class TrackerManager:
    DUMMY = Person("--- DUMMY ---", "--- DUMMY ---", ["--- DUMMY ---"])

    def __init__(self, add_dummy=False):
        self.people: dict[Person, list[EmailMessage]] = {}
        self.email_mappings = {}  # {recv_email: map_email}

        # Adds a dummy email if we want to count ALL emails, not just matched emails
        self.add_dummy = add_dummy
        if add_dummy:
            self.people[TrackerManager.DUMMY] = []

        self.blacklist = set()

    def load_email_blacklist(self, blacklist_file: str):
        if not os.path.exists(blacklist_file):
            print("Could not find blacklist email file")
            return
        elif not blacklist_file.endswith(".txt"):
            print("Inputted blacklist file is not a .txt file")
            return

        with open(blacklist_file, encoding="latin-1") as f:
            for i in f:
                self.blacklist.add(i.strip().lower().replace(".", ""))

    def load_tracker_csv(self, tracker_export_file: str):
        if not os.path.exists(tracker_export_file):
            print("Could not find tracker file")
            return
        elif not tracker_export_file.endswith(".csv"):
            print("Inputted tacker file is not a .csv file")
            return

        with open(tracker_export_file, encoding="latin-1") as f:
            reader = csv.reader(f)
            next(reader)

            for row in reader:
                # Blank Lines - Everyone should have a last name
                if not row[2].strip():
                    continue

                emails = [row[3]]
                for i in row[8].split("\n"):
                    if match := re.search(r"[^\s@]+@[^\s@]+", i):
                        emails.append(match.group(0))

                self.people[Person(f"{row[0]} {row[2]}", f"{row[1]} {row[2]}", emails)] = []

        self.update_email_mapping()

    def load_email_mapping(self, email_map_file):
        if not os.path.exists(email_map_file):
            print("Could not find map file")
            return
        elif not email_map_file.endswith(".json"):
            print("Inputted map file is not a .json file")
            return

        with open(email_map_file, encoding="latin-1") as f:
            self.email_mappings.update({i: j["map_email"] for i, j in json.load(f).items() if j["map_email"]})

        self.update_email_mapping()

    def update_email_mapping(self):
        for recv, data in self.email_mappings.items():
            for p in self.people:
                if data in p.emails:
                    p.add_email(recv)

    def _find_matching_person(self, msg: EmailMessage):
        email_matches = []
        name_matches = []

        for p in self.people:
            if msg.does_email_match(p):
                email_matches.append(p)

            if msg.does_name_match(p):
                name_matches.append(p)

        return email_matches, name_matches

    def _email_in_blacklist(self, person: Person):
        return not self.blacklist.isdisjoint(person.sanitized_emails)

    def compile_emails(self, email_export: list[EmailMessage]):
        unknown_emails = []

        for e in email_export:
            if self._email_in_blacklist(e):
                continue
            elif not e.emails:
                print(f"Could not find email for `{e.name}`, skipping")
                continue

            email_matches, name_matches = self._find_matching_person(e)
            unknown = True

            if email_matches:
                if len(email_matches) == 1:
                    self.people[email_matches[0]].append(e)
                    unknown = False
                else:
                    print(f"Message {e} email got matched with multiple people: {email_matches}, skipping")
            elif name_matches:
                if len(name_matches) == 1:
                    self.people[name_matches[0]].append(e)
                    unknown = False
                else:
                    print(f"Message {e} name got matched with multiple people: {name_matches}, skipping")

            if unknown:
                unknown_emails.append(e)

                if self.add_dummy:
                    self.people[TrackerManager.DUMMY].append(e)

        return {(e.name, e.emails.pop() if e.emails else None) for e in unknown_emails}

    @staticmethod
    def generate_mapping(unknown):
        return {i[1]: {"name": i[0], "map_email": ""} for i in unknown if i[1]}

    def extract_weekly_emails(self, start_time: datetime.date, num_weeks):
        start_dt = int(datetime.datetime.combine(start_time, datetime.time(), tzinfo=datetime.timezone.utc).timestamp() * 1000)
        if start_time.weekday() != 0:
            start_dt += (7 - start_time.weekday()) * 86400000

        email_counts = [0 for _ in range(num_weeks)]

        for emails in self.people.values():
            for e in emails:
                weeks_after = (e.receive_time - start_dt) // 604800000
                if weeks_after >= num_weeks:
                    # Don't count this email since it's weird
                    print(f"Email did not found within bounds: {weeks_after}")
                    continue

                if weeks_after < 0:
                    # Just saying everything before the 'start time' is just first week
                    weeks_after = 0

                email_counts[weeks_after] += 1

        return email_counts


def export_mapping(unknown, mapping_file=None):
    mapping = TrackerManager.generate_mapping(unknown)

    if mapping_file and os.path.exists(mapping_file):
        with open(mapping_file) as f:
            mapping.update(json.load(f))

    with open(mapping_file or "data/email_mappings.json", "w") as f:
        json.dump(mapping, f, indent=2)
