import json
import re
import os
import csv

from libratom.lib.pff import PffArchive
import pypff
from bs4 import BeautifulSoup


class Person:
    def __init__(self, name: str, preferred_name: str, emails: list[str]):
        self.name = (name or "").strip()
        self.preferred_name = (preferred_name or "").strip()
        self.emails = {i.strip() for i in emails}

    def __repr__(self):
        return f"{self.__class__.__name__}(name={self.name}, preferred_name={self.preferred_name}, emails={self.emails})"

    def _get_sanitized_names(self):
        return re.sub(r"[^a-z\s]", "", self.name.lower()).split(), re.sub(r"[^a-z\s]", "",
                                                                          self.preferred_name.lower()).split()

    def is_same_person(self, other: 'Person'):
        if self.does_email_match(other):
            return True

        return self.does_name_match(other)

    def does_email_match(self, other: 'Person'):
        return not self.emails.isdisjoint(other.emails)

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
        match = re.search(r"From: (.+? )?<(.+?)>", message.transport_headers or "")
        if match:
            sender_email = [match.group(2)]
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
        # Jank way of cutting things off at the first <hr>
        parsed_html = BeautifulSoup((html or b"").decode("latin-1").split("<hr")[0], "html.parser")

        # get_text() has some weird new lines
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
    def __init__(self):
        self.people: dict[Person, list[EmailMessage]] = {}
        self.email_mappings = {}  # {recv_email: map_email}

    def load_tracker_csv(self, tracker_export_file: str):
        if not os.path.exists(tracker_export_file):
            print("Could not find tracker file")
            return
        elif not tracker_export_file.endswith(".csv"):
            print("Inputted tacker file is not a .csv file")
            return

        with open(tracker_export_file, encoding="utf8") as f:
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

        with open(email_map_file, encoding="utf8") as f:
            self.email_mappings.update({i: j["map_email"] for i, j in json.load(f).items() if j["map_email"]})

        self.update_email_mapping()

    def update_email_mapping(self):
        for recv, track in self.email_mappings.items():
            for p in self.people:
                if track in p.emails:
                    p.emails.add(recv)

    def _find_matching_person(self, msg: EmailMessage):
        email_matches = []
        name_matches = []

        for p in self.people:
            if msg.does_email_match(p):
                email_matches.append(p)

            if msg.does_name_match(p):
                name_matches.append(p)

        return email_matches, name_matches

    def compile_emails(self, email_export: list[EmailMessage]):
        unknown = []

        for e in email_export:
            email_matches, name_matches = self._find_matching_person(e)
            if email_matches:
                if len(email_matches) == 1:
                    self.people[email_matches[0]].append(e)
                else:
                    print(f"Message {e} email got matched with multiple people: {email_matches}, skipping")
                    unknown.append(e)
            elif name_matches:
                if len(name_matches) == 1:
                    self.people[name_matches[0]].append(e)
                else:
                    print(f"Message {e} name got matched with multiple people: {name_matches}, skipping")
                    unknown.append(e)
            else:
                unknown.append(e)

        return {(e.name, e.emails.pop() if e.emails else None) for e in unknown}

    @staticmethod
    def generate_mapping(unknown):
        return {i[1]: {"name": i[0], "map_email": ""} for i in unknown}
