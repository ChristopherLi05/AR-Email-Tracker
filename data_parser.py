import re
import os
import csv

from libratom.lib.pff import PffArchive
import pypff


class Person:
    def __init__(self, name: str, preferred_name: str, emails: list[str]):
        self.name = (name or "").strip()
        self.preferred_name = (preferred_name or "").strip()
        self.emails = {i.strip() for i in emails}

    def __repr__(self):
        return f"{self.__class__.__name__}(name={self.name}, preferred_name={self.preferred_name}, emails={self.emails})"

    def _get_sanitized_names(self):
        return re.sub(r"[^a-z\s]", "", self.name.lower()).split(), re.sub(r"[^a-z\s]", "", self.preferred_name.lower()).split()

    def is_same_person(self, other: 'Person'):
        # Basic Email Matching
        if not self.emails.isdisjoint(other.emails):
            return True

        # Basic Name Matching ; We know that there's at least 1 intersection if the length < 4
        # if not {self.name, self.preferred_name}.isdisjoint({other.name, other.preferred_name}):
        #     return True
        if self.name == other.name or self.preferred_name == other.preferred_name:
            return True

        for name1, name2 in zip(self._get_sanitized_names(), other._get_sanitized_names()):
            if name1 == name2 and name1:
                return True
            # First + Last Name matching
            # if name1[0] == name2[0] and name1[-1] == name2[-1]:
            #     return True

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

    def __repr__(self):
        return f"{self.__class__.__name__}(name={self.name}, preferred_name={self.preferred_name}, emails={self.emails}, receive_time={self.receive_time})"


class EmailCompiler:
    def __init__(self):
        self.emails = {}

    def add_person(self, other: EmailMessage):
        for e in self.emails:
            if e.is_same_person(other):
                self.emails[e] += 1
                break
        else:
            self.emails[other] = 1


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


def extract_tracker(tracker_export_file: str):
    if not os.path.exists(tracker_export_file):
        print("Could not find file")
        return []
    elif not tracker_export_file.endswith(".csv"):
        print("Inputted file is not a .csv file")
        return []

    people = []
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

            people.append(Person(f"{row[0]} {row[2]}", f"{row[1]} {row[2]}", emails))
    return people


def compile_emails(extracted_email_data: list[EmailMessage]):
    emails = {}

    for msg in extracted_email_data:
        for e in emails:
            if emails[e][0].is_same_person(msg):
                emails[e].append(msg)
                break
        else:
            emails[msg.name] = [msg]

    return emails


if __name__ == "__main__":
    # print(extract_emails("data/email_export.pst"))
    data = extract_emails("data/email_export.pst")
    data1 = compile_emails(data)

    for i, j in data1.items():
        print(i, len(j))

    # print(extract_tracker("data/tracker_export.csv"))
    # data = extract_tracker("data/tracker_export.csv")
    # for i in range(len(data)):
    #     for j in range(i + 1, len(data)):
    #         if data[i].is_same_person(data[j]):
    #             print(data[i], data[j])
    #             input()
