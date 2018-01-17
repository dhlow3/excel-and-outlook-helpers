# -*- coding: utf-8 -*-
"""Module for exporting attachments from Outlook desktop Application."""
import os
import win32com.client

from datetime import datetime, timedelta


class Export:
    """Export base class."""

    def __init__(self, dst=''):
        """Initialize an export.

        Parameters
        ----------
        dst: str
            Location to place exported files (Default: Downloads folder).
        """
        if not dst:
            # Find path to user's Downloads folder
            oShell = win32com.client.Dispatch("Wscript.Shell")
            downloads = oShell.RegRead(
                'HKEY_CURRENT_USER\\Software\Microsoft\\Windows'
                '\\CurrentVersion\\Explorer\\User Shell Folders'
                '\\{374DE290-123F-4565-9164-39C4925E467B}')
            assert os.path.exists(downloads), 'downloads path does not exist'
            self.dst = downloads
        else:
            assert os.path.exists(dst), 'destination path does not exist'
            self.dst = dst

        self.exported_files = []

    def rename_file(self, file_index, new_name):
        """Rename a file in self.exported_files.

        Parameters
        ----------
        file_index: int
            The index of the file in self.exported_files to be renamed.
        new_name: str
            The new name of the file.
        """
        assert isinstance(file_index, int)
        assert isinstance(new_name, str)

        exported_file = self.exported_files[file_index]
        # Separate folder path from filename
        head, tail = os.path.split(exported_file)
        # Get the current file extension
        name, extension = os.path.splitext(tail)
        # Create path for new filename
        if '.' in new_name:
            new_file = os.path.join(head, new_name)
        else:
            new_file = os.path.join(head, new_name + extension)
        # Rename file
        os.rename(exported_file, new_file)
        # Drop old file from self.exported_files
        self.exported_files.pop(file_index)
        # Append new file to self.exported_files
        self.exported_files.append(new_file)


class EmailExport(Export):
    """A class for email attachment exports."""

    def export_email_attachment(self, subject, lookback=0,
                                inbox_subfolder=None,
                                date=None):
        """Export an attachment with a specific subject and lookback.

        Parameters
        ----------
        subject: str
            A keyword to find in an email subject line.
        lookback: int
            The number of days to look back for an email (Default: 0 = today).
        inbox_subfolder: str
            The name of an Email inbox subfolder (Default: None).
        date: tuple
            A date used for filtering emails (Format: month, day, 4-d year).

        Returns
        ----------
        exported_files: list
            The exported attachment paths.
        """
        # Parse email inbox
        outlook = win32com.client.Dispatch(
            "Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # Go to the inbox folder "6"

        # Specify a subfolder to search
        if inbox_subfolder:
            messages = inbox.Folders.Item(inbox_subfolder).Items
        else:
            messages = inbox.Items  # Get all the messages

        if lookback:
            # Search emails x days ago
            date = datetime.now() - timedelta(days=lookback)
            start = datetime(date.year, date.month, date.day)\
                .strftime('%m/%d/%Y 12:00AM')
            end = datetime(date.year, date.month, date.day)\
                .strftime('%m/%d/%Y 11:59PM')
            date_filter = ("[LastModificationTime] >= \'{}\' AND "
                           "[LastModificationTime] <= \'{}\'"
                           .format(start, end))
        elif date:
            # Search emails on a specified date
            start = datetime(date[-1], date[0], date[1])\
                .strftime('%m/%d/%Y 12:00AM')
            end = datetime(date[-1], date[0], date[1])\
                .strftime('%m/%d/%Y 11:59PM')
            date_filter = ("[LastModificationTime] >= \'{}\' AND "
                           "[LastModificationTime] <= \'{}\'"
                           .format(start, end))
        else:
            # If no lookback or date is specified search current day emails
            date_filter = "@SQL=%today(DAV:getlastmodified)%"

        # Emails between start and end dates
        inbox_by_date = messages.Restrict(date_filter)

        # Search email subjects for a specified keyword
        subject_tag = ('http://schemas.microsoft.com/mapi/proptag/0x0037001E '
                       "ci_phrasematch '{}'".format(subject))
        subject_filter = '@SQL={}'.format(subject_tag)
        inbox_subject = inbox_by_date.Restrict(subject_filter)

        # Download attachment
        message = inbox_subject.GetLast()  # Get last email from results
        attachments = message.Attachments

        # Download all attachments in email
        for attachment in attachments:
            attachment_path = os.path.join(self.dst, attachment.FileName)
            attachment.SaveAsFile(attachment_path)
            self.exported_files.append(attachment_path)

        return self.exported_files
