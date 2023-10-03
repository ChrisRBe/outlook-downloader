"""
Copyright: ChrisRBe 2023-10-03
"""
import sys
import win32com.client


from pathlib import Path


def main():
    """
    connect to outlook and save attachments

    :return:
    """
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # TODO: Pulling this one out into a command line parameter might make more sense
    storage = Path.cwd()

    # TODO: The tool should probably include something like additional sub commands to get this required
    # TODO: information. Setting the folder name should be passed in as a command line parameter as well.
    what = outlook.Folders.Item(1)
    something = what.Folders["Telekom"]

    emails = something.Items

    for email in emails:
        print(email)
        email_received = email.ReceivedTime

        for attachment in email.Attachments:
            attachment_file_name = attachment.FileName
            # TODO: does this make sense here as is or is there a better way to select the relevant files/file types?
            if attachment_file_name.endswith("zip") or attachment_file_name.endswith("pdf"):
                attachment_filename = f"{storage}/{email_received:%Y%m%d}_{attachment_file_name}"
                attachment.SaveAsFile(attachment_filename)


if __name__ == "__main__":
    sys.exit(main())
