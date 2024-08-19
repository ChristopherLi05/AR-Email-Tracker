import data_parser


def main():
    manager = data_parser.TrackerManager()
    manager.load_tracker_csv("data/mps_tracker_export.csv")
    manager.load_email_mapping("data/email_mappings.json")
    manager.load_email_blacklist("data/blacklist.txt")

    data = data_parser.extract_emails("data/email_exports/chris_email_export.pst")

    unknown_emails = manager.compile_emails(data)

    print(unknown_emails)
    print(len(unknown_emails))

    data_parser.export_mapping(unknown_emails, "data/email_mappings.json")

    # unknown = manager.compile_emails(data)
    # emails = set()
    #
    # for i in unknown:
    #     emails.update(i.emails)
    #
    # print(emails)

    # data1 = compile_emails(data)

    # for i, j in data1.items():
    #     print(i, len(j))

    # print(extract_tracker("data/mps_tracker_export.csv"))
    # data = extract_tracker("data/mps_tracker_export.csv")
    # for i in range(len(data)):
    #     for j in range(i + 1, len(data)):
    #         if data[i].is_same_person(data[j]):
    #             print(data[i], data[j])
    #             input()


if __name__ == "__main__":
    main()
