"""
Microsoft Outlook COM client.

Provides a high-level Python interface to automate Microsoft Outlook
using the COM automation API (Outlook.Application).
"""

from __future__ import annotations

import csv
import datetime
import os
from typing import Any

import pythoncom
import win32com.client


class OutlookClient:
    """Client for interacting with Microsoft Outlook via COM automation."""

    _PROG_IDS = [
        "Outlook.Application",
    ]

    # Outlook folder constants
    OL_FOLDER_INBOX = 6
    OL_FOLDER_SENT = 5
    OL_FOLDER_DRAFTS = 16
    OL_FOLDER_DELETED = 3
    OL_FOLDER_CALENDAR = 9
    OL_FOLDER_CONTACTS = 10
    OL_FOLDER_TASKS = 13
    OL_FOLDER_NOTES = 12
    OL_FOLDER_JOURNAL = 11
    OL_FOLDER_OUTBOX = 4

    # Mail importance
    IMPORTANCE_LOW = 0
    IMPORTANCE_NORMAL = 1
    IMPORTANCE_HIGH = 2

    def __init__(self) -> None:
        """Initialize the Outlook client and connect to Outlook."""
        self._app: Any = None
        self._namespace: Any = None
        self._connect()

    def _connect(self) -> None:
        """Connect to a running Outlook instance or create a new one."""
        pythoncom.CoInitialize()

        for prog_id in self._PROG_IDS:
            # First try getting an already running instance
            try:
                self._app = win32com.client.GetActiveObject(prog_id)
            except Exception:
                pass

            # If no running instance, create a new one
            if self._app is None:
                try:
                    self._app = win32com.client.Dispatch(prog_id)
                except Exception:
                    continue

            # Verify the connection works by getting the MAPI namespace
            if self._app is not None:
                try:
                    self._namespace = self._app.GetNamespace("MAPI")
                    break
                except Exception:
                    self._app = None
                    continue

        if self._app is None:
            raise RuntimeError(
                "Could not connect to Microsoft Outlook. "
                "Please ensure Microsoft Outlook is installed and running."
            )

    @property
    def app(self) -> Any:
        """Get the underlying COM application object."""
        if self._app is None:
            self._connect()
        return self._app

    @property
    def namespace(self) -> Any:
        """Get the MAPI namespace."""
        if self._namespace is None:
            self._connect()
        return self._namespace

    def _get_folder(self, folder_type: int, account_name: str | None = None) -> Any:
        """Get a default folder by type."""
        if account_name:
            # Find the specific account
            for acc in self.namespace.Accounts:
                if acc.DisplayName.lower() == account_name.lower():
                    # The default store for the account
                    store = acc.DeliveryStore
                    return store.GetDefaultFolder(folder_type)
        return self.namespace.GetDefaultFolder(folder_type)

    # ── Mail Operations ─────────────────────────────────────────────

    def list_emails(
        self,
        folder_type: int = OL_FOLDER_INBOX,
        count: int = 20,
        offset: int = 0,
        fields: str | None = None,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List recent emails from a folder.

        Args:
            folder_type: The Outlook folder constant (default: Inbox).
            count: Maximum number of emails to return (default: 20).
            offset: Number of emails to skip (for pagination).
            fields: Comma-separated field names to include (e.g. "subject,sender_name").
                    When omitted, returns a compact summary.
            account_name: Optional account display name to target a specific account.

        Returns:
            List of compact email summary dicts.
        """
        folder = self._get_folder(folder_type, account_name)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort descending

        result = []
        start = 1 + offset
        end = min(start + count, items.Count + 1)
        for i in range(start, end):
            try:
                item = items.Item(i)
                # Only process MailItem (class 43)
                if item.Class == 43:
                    d = self._mail_to_summary(item)
                    result.append(self._filter_fields(d, fields))
            except Exception:
                continue

        return result

    def search_emails(
        self,
        query: str = "",
        folder_type: int = OL_FOLDER_INBOX,
        count: int = 20,
        offset: int = 0,
        fields: str | None = None,
        subject: str | None = None,
        sender: str | None = None,
        received_after: str | None = None,
        received_before: str | None = None,
        unread_only: bool = False,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        Search emails with various filters.

        Args:
            query: DASL filter string (advanced). If empty, builds from other filters.
            folder_type: Folder to search in (default: Inbox).
            count: Maximum results (default: 20).
            offset: Number of results to skip (for pagination).
            fields: Comma-separated field names to include. When omitted, returns a compact summary.
            subject: Filter by subject containing this text.
            sender: Filter by sender name/email containing this text.
            received_after: ISO date string (e.g., '2026-01-01').
            received_before: ISO date string.
            unread_only: Only return unread emails.
            account_name: Optional account display name.

        Returns:
            List of matching compact email summary dicts.
        """
        folder = self._get_folder(folder_type, account_name)
        items = folder.Items

        # Build DASL filter if not explicitly provided
        if not query:
            filters = []
            if subject:
                filters.append(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject}%'")
            if sender:
                filters.append(
                    f"@SQL=\"urn:schemas:httpmail:sendername\" LIKE '%{sender}%'"
                )
            if received_after:
                filters.append(
                    f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{received_after}'"
                )
            if received_before:
                filters.append(
                    f"@SQL=\"urn:schemas:httpmail:datereceived\" <= '{received_before}'"
                )
            if unread_only:
                filters.append("@SQL=\"urn:schemas:httpmail:read\" = 0")

            query = " AND ".join(filters) if filters else ""

        try:
            if query:
                filtered = items.Restrict(query)
            else:
                filtered = items
        except Exception:
            filtered = items

        filtered.Sort("[ReceivedTime]", True)

        result = []
        start = 1 + offset
        end = min(start + count, filtered.Count + 1)
        for i in range(start, end):
            try:
                item = filtered.Item(i)
                if item.Class == 43:
                    d = self._mail_to_summary(item)
                    result.append(self._filter_fields(d, fields))
            except Exception:
                continue

        return result

    def send_email(
        self,
        to: str,
        subject: str = "",
        body: str = "",
        cc: str = "",
        bcc: str = "",
        attachments: list[str] | None = None,
        html_body: bool = False,
        importance: int = IMPORTANCE_NORMAL,
    ) -> dict[str, Any]:
        """
        Send a new email.

        Args:
            to: Recipient email(s), semicolon-separated.
            subject: Email subject.
            body: Email body text.
            cc: CC recipients.
            bcc: BCC recipients.
            attachments: List of file paths to attach.
            html_body: If True, treat body as HTML.
            importance: Importance level (0=Low, 1=Normal, 2=High).

        Returns:
            Dict with sent email info.
        """
        mail = self.app.CreateItem(0)  # 0 = olMailItem
        mail.To = to
        mail.Subject = subject
        if html_body:
            mail.HTMLBody = body
        else:
            mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Importance = importance

        if attachments:
            for filepath in attachments:
                full_path = os.path.abspath(filepath)
                if os.path.exists(full_path):
                    mail.Attachments.Add(full_path)

        mail.Send()

        return {
            "message": "Email sent successfully.",
            "to": to,
            "subject": subject,
        }

    def create_draft(
        self,
        subject: str = "",
        body: str = "",
        to: str = "",
        cc: str = "",
        bcc: str = "",
        attachments: list[str] | None = None,
        html_body: bool = False,
        importance: int = IMPORTANCE_NORMAL,
    ) -> dict[str, Any]:
        """
        Create a draft email (does NOT send).

        Args:
            subject: Email subject.
            body: Email body text.
            to: Recipient email(s), semicolon-separated.
            cc: CC recipients.
            bcc: BCC recipients.
            attachments: List of file paths to attach.
            html_body: If True, treat body as HTML.
            importance: Importance level (0=Low, 1=Normal, 2=High).

        Returns:
            Dict with draft info including EntryID.
        """
        mail = self.app.CreateItem(0)  # 0 = olMailItem
        if to:
            mail.To = to
        mail.Subject = subject
        if html_body:
            mail.HTMLBody = body
        else:
            mail.Body = body
        if cc:
            mail.CC = cc
        if bcc:
            mail.BCC = bcc
        mail.Importance = importance

        if attachments:
            for filepath in attachments:
                full_path = os.path.abspath(filepath)
                if os.path.exists(full_path):
                    mail.Attachments.Add(full_path)

        # Save to Drafts folder instead of sending
        mail.Save()

        return {
            "message": "Draft saved successfully.",
            "entry_id": mail.EntryID,
            "to": to,
            "subject": subject,
        }

    def reply_email(
        self,
        entry_id: str,
        body: str = "",
        reply_all: bool = False,
        html_body: bool = False,
        attachments: list[str] | None = None,
    ) -> dict[str, Any]:
        """
        Reply to an email identified by its EntryID.

        Args:
            entry_id: The Outlook EntryID of the email to reply to.
            body: Reply body text.
            reply_all: If True, use ReplyAll instead of Reply.
            html_body: If True, treat body as HTML.
            attachments: List of file paths to attach.

        Returns:
            Dict with reply info.
        """
        original = self.namespace.GetItemFromID(entry_id)
        if reply_all:
            reply = original.ReplyAll()
        else:
            reply = original.Reply()

        if html_body:
            reply.HTMLBody = body + reply.HTMLBody if body else reply.HTMLBody
        elif body:
            reply.Body = body + "\n\n" + reply.Body

        if attachments:
            for filepath in attachments:
                full_path = os.path.abspath(filepath)
                if os.path.exists(full_path):
                    reply.Attachments.Add(full_path)

        reply.Send()

        return {
            "message": "Reply sent successfully.",
            "original_subject": original.Subject,
        }

    def forward_email(
        self,
        entry_id: str,
        to: str,
        body: str = "",
        html_body: bool = False,
        attachments: list[str] | None = None,
    ) -> dict[str, Any]:
        """
        Forward an email to new recipients.

        Args:
            entry_id: The Outlook EntryID of the email to forward.
            to: Recipient email(s).
            body: Additional body text.
            html_body: If True, treat body as HTML.
            attachments: List of file paths to attach.

        Returns:
            Dict with forward info.
        """
        original = self.namespace.GetItemFromID(entry_id)
        forward = original.Forward()
        forward.To = to

        if html_body:
            forward.HTMLBody = body + forward.HTMLBody if body else forward.HTMLBody
        elif body:
            forward.Body = body + "\n\n" + forward.Body

        if attachments:
            for filepath in attachments:
                full_path = os.path.abspath(filepath)
                if os.path.exists(full_path):
                    forward.Attachments.Add(full_path)

        forward.Send()

        return {
            "message": "Email forwarded successfully.",
            "to": to,
            "original_subject": original.Subject,
        }

    def delete_email(self, entry_id: str) -> dict[str, Any]:
        """
        Delete an email by EntryID.

        Args:
            entry_id: The Outlook EntryID of the email.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        subject = item.Subject
        item.Delete()
        return {"message": f"Deleted email: '{subject}'"}

    def move_email(self, entry_id: str, dest_folder_type: int) -> dict[str, Any]:
        """
        Move an email to a different folder.

        Args:
            entry_id: The Outlook EntryID of the email.
            dest_folder_type: The Outlook folder constant for destination.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        dest = self._get_folder(dest_folder_type)
        subject = item.Subject
        item.Move(dest)
        return {"message": f"Moved email '{subject}' to target folder."}

    def mark_as_read(self, entry_id: str, read: bool = True) -> dict[str, Any]:
        """
        Mark an email as read or unread.

        Args:
            entry_id: The Outlook EntryID of the email.
            read: True to mark as read, False to mark as unread.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        item.UnRead = not read
        status = "read" if read else "unread"
        return {"message": f"Marked email '{item.Subject}' as {status}."}

    def get_email_by_id(self, entry_id: str) -> dict[str, Any]:
        """
        Get full details of an email by its EntryID.

        Args:
            entry_id: The Outlook EntryID.

        Returns:
            Dict with full email details.
        """
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 43:
            raise ValueError(f"Item is not an email (class={item.Class}).")
        return self._mail_to_dict(item, include_body=True)

    def save_attachment(
        self,
        entry_id: str,
        attachment_index: int = 1,
        save_path: str = "",
    ) -> dict[str, Any]:
        """
        Save an attachment from an email to disk.

        Args:
            entry_id: The Outlook EntryID of the email.
            attachment_index: 1-based index of attachment.
            save_path: Directory or full file path to save to.

        Returns:
            Dict with saved file info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 43:
            raise ValueError("Item is not an email.")

        if item.Attachments.Count < attachment_index:
            raise IndexError(
                f"Attachment index {attachment_index} out of range "
                f"(total: {item.Attachments.Count})."
            )

        att = item.Attachments.Item(attachment_index)
        if save_path:
            if os.path.isdir(save_path):
                full_path = os.path.join(save_path, att.FileName)
            else:
                full_path = save_path
        else:
            full_path = os.path.join(os.getcwd(), att.FileName)

        att.SaveAsFile(full_path)
        return {"message": f"Saved attachment to: {full_path}", "file": full_path}

    # ── Calendar Operations ──────────────────────────────────────────

    def list_calendar_events(
        self,
        start_date: str | None = None,
        end_date: str | None = None,
        count: int = 20,
        offset: int = 0,
        fields: str | None = None,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List calendar events.

        Args:
            start_date: ISO date string for range start (default: today).
            end_date: ISO date string for range end (default: +30 days).
            count: Maximum events to return (default: 20).
            offset: Number of events to skip (for pagination).
            fields: Comma-separated field names to include. When omitted, returns a compact summary.
            account_name: Optional account display name.

        Returns:
            List of compact event summary dicts (no body).
        """
        folder = self._get_folder(self.OL_FOLDER_CALENDAR, account_name)
        items = folder.Items
        items.Sort("[Start]")
        items.IncludeRecurrences = True

        # Build date filter
        if not start_date:
            start_date = datetime.date.today().isoformat()
        if not end_date:
            end_date = (datetime.date.today() + datetime.timedelta(days=30)).isoformat()

        filter_str = (
            f"@SQL=\"urn:schemas:calendar:dtstart\" >= '{start_date}' AND "
            f"\"urn:schemas:calendar:dtend\" <= '{end_date}'"
        )

        try:
            filtered = items.Restrict(filter_str)
        except Exception:
            filtered = items

        result = []
        start = 1 + offset
        end = min(start + count, filtered.Count + 1)
        for i in range(start, end):
            try:
                item = filtered.Item(i)
                if item.Class == 26:  # olAppointment
                    d = self._appointment_to_summary(item)
                    result.append(self._filter_fields(d, fields))
            except Exception:
                continue

        return result

    def create_appointment(
        self,
        subject: str,
        start_time: str,
        end_time: str,
        body: str = "",
        location: str = "",
        all_day: bool = False,
        reminder_minutes: int = 15,
        recipients: str = "",
    ) -> dict[str, Any]:
        """
        Create a new calendar appointment.

        Args:
            subject: Appointment subject.
            start_time: ISO datetime string for start.
            end_time: ISO datetime string for end.
            body: Appointment body/notes.
            location: Location string.
            all_day: If True, create an all-day event.
            reminder_minutes: Minutes before event to remind (0 for none).
            recipients: Optional attendees (semicolon-separated email addresses).

        Returns:
            Dict with created appointment info.
        """
        apt = self.app.CreateItem(1)  # 1 = olAppointmentItem
        apt.Subject = subject

        # Convert ISO strings to datetime objects for reliable COM compatibility
        try:
            apt.Start = datetime.datetime.fromisoformat(start_time)
            apt.End = datetime.datetime.fromisoformat(end_time)
        except (ValueError, TypeError):
            # Fallback: try string assignment
            apt.Start = start_time
            apt.End = end_time

        apt.Body = body
        apt.Location = location
        apt.AllDayEvent = all_day
        apt.ReminderMinutesBeforeStart = reminder_minutes
        apt.ReminderSet = reminder_minutes > 0

        if recipients:
            apt.RequiredAttendees = recipients
            # Send invitation
            apt.MeetingStatus = 1  # olMeeting
            apt.Send()
        else:
            apt.Save()

        return {
            "message": "Appointment created successfully.",
            "entry_id": apt.EntryID,
            "subject": subject,
            "start": start_time,
            "end": end_time,
            "location": location,
        }

    def delete_appointment(self, entry_id: str) -> dict[str, Any]:
        """
        Delete a calendar appointment by EntryID.

        Args:
            entry_id: The Outlook EntryID of the appointment.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        subject = item.Subject
        item.Delete()
        return {"message": f"Deleted appointment: '{subject}'"}

    # ── Contacts Operations ──────────────────────────────────────────

    def list_contacts(
        self,
        count: int = 50,
        offset: int = 0,
        fields: str | None = None,
        search: str = "",
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List contacts.

        Args:
            count: Maximum contacts to return (default: 50).
            offset: Number of contacts to skip (for pagination).
            fields: Comma-separated field names to include. When omitted, returns a compact summary.
            search: Optional text to filter contacts (searches name and email).
            account_name: Optional account display name.

        Returns:
            List of compact contact summary dicts.
        """
        folder = self._get_folder(self.OL_FOLDER_CONTACTS, account_name)
        items = folder.Items

        if search:
            filter_str = (
                f"@SQL=\"urn:schemas:contacts:cn\" LIKE '%{search}%' OR "
                f"\"urn:schemas:contacts:email1address\" LIKE '%{search}%'"
            )
            try:
                items = items.Restrict(filter_str)
            except Exception:
                pass

        items.Sort("[FullName]")

        result = []
        start = 1 + offset
        end = min(start + count, items.Count + 1)
        for i in range(start, end):
            try:
                item = items.Item(i)
                if item.Class == 40:  # olContact
                    d = self._contact_to_summary(item)
                    result.append(self._filter_fields(d, fields))
            except Exception:
                continue

        return result

    def get_contact_by_id(self, entry_id: str) -> dict[str, Any]:
        """Get full contact details by EntryID."""
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 40:
            raise ValueError(f"Item is not a contact (class={item.Class}).")
        return self._contact_to_dict(item)

    def create_contact(
        self,
        full_name: str,
        email: str = "",
        phone: str = "",
        mobile: str = "",
        company: str = "",
        job_title: str = "",
    ) -> dict[str, Any]:
        """
        Create a new contact.

        Args:
            full_name: Contact's full name.
            email: Email address.
            phone: Business phone number.
            mobile: Mobile phone number.
            company: Company name.
            job_title: Job title.

        Returns:
            Dict with created contact info.
        """
        contact = self.app.CreateItem(2)  # 2 = olContactItem
        contact.FullName = full_name
        if email:
            contact.Email1Address = email
        if phone:
            contact.BusinessTelephoneNumber = phone
        if mobile:
            contact.MobileTelephoneNumber = mobile
        if company:
            contact.CompanyName = company
        if job_title:
            contact.JobTitle = job_title

        contact.Save()

        return {
            "message": "Contact created successfully.",
            "full_name": full_name,
            "email": email,
        }

    def delete_contact(self, entry_id: str) -> dict[str, Any]:
        """
        Delete a contact by EntryID.

        Args:
            entry_id: The Outlook EntryID of the contact.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        name = item.FullName
        item.Delete()
        return {"message": f"Deleted contact: '{name}'"}

    # ── Mailbox Info ─────────────────────────────────────────────────

    def get_mailbox_info(self) -> dict[str, Any]:
        """
        Get information about the Outlook mailbox.

        Returns:
            Dict with account names, folder counts, etc.
        """
        accounts = []
        for acc in self.namespace.Accounts:
            accounts.append({
                "display_name": str(acc.DisplayName),
                "smtp_address": str(acc.SmtpAddress),
                "account_type": str(acc.AccountType),
            })

        # Get unread counts for common folders
        folder_info = {}
        folder_names = {
            self.OL_FOLDER_INBOX: "Inbox",
            self.OL_FOLDER_SENT: "Sent",
            self.OL_FOLDER_DRAFTS: "Drafts",
            self.OL_FOLDER_DELETED: "Deleted",
            self.OL_FOLDER_CALENDAR: "Calendar",
            self.OL_FOLDER_CONTACTS: "Contacts",
        }

        for folder_type, name in folder_names.items():
            try:
                folder = self._get_folder(folder_type)
                unread = folder.UnReadItemCount if hasattr(folder, "UnReadItemCount") else 0
                total = folder.Items.Count
                folder_info[name] = {"total": total, "unread": unread}
            except Exception:
                folder_info[name] = {"total": "N/A", "unread": "N/A"}

        return {
            "version": str(self.app.Version),
            "accounts": accounts,
            "folders": folder_info,
        }

    # ── Calendar: Detail / Update / Respond ──────────────────────────

    def get_appointment_by_id(self, entry_id: str) -> dict[str, Any]:
        """Get full details of a calendar appointment by its EntryID."""
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 26:  # olAppointment
            raise ValueError(f"Item is not an appointment (class={item.Class}).")
        return self._appointment_to_dict(item)

    def update_appointment(
        self,
        entry_id: str,
        subject: str | None = None,
        start_time: str | None = None,
        end_time: str | None = None,
        body: str | None = None,
        location: str | None = None,
        all_day: bool | None = None,
        reminder_minutes: int | None = None,
        recipients: str | None = None,
        send_update: bool = False,
    ) -> dict[str, Any]:
        """Update an existing calendar appointment. Only provided fields are changed.

        Args:
            entry_id: The Outlook EntryID of the appointment.
            subject: New subject.
            start_time: New start time (ISO datetime string).
            end_time: New end time (ISO datetime string).
            body: New body/notes.
            location: New location.
            all_day: Set all-day event.
            reminder_minutes: Minutes before to remind (0 for none).
            recipients: Updated attendees (semicolon-separated emails).
            send_update: If True, send an update to attendees.

        Returns:
            Dict with update result.
        """
        apt = self.namespace.GetItemFromID(entry_id)
        if apt.Class != 26:
            raise ValueError(f"Item is not an appointment (class={apt.Class}).")

        changed = []
        if subject is not None:
            apt.Subject = subject
            changed.append("subject")
        if start_time is not None:
            try:
                apt.Start = datetime.datetime.fromisoformat(start_time)
            except (ValueError, TypeError):
                apt.Start = start_time
            changed.append("start_time")
        if end_time is not None:
            try:
                apt.End = datetime.datetime.fromisoformat(end_time)
            except (ValueError, TypeError):
                apt.End = end_time
            changed.append("end_time")
        if body is not None:
            apt.Body = body
            changed.append("body")
        if location is not None:
            apt.Location = location
            changed.append("location")
        if all_day is not None:
            apt.AllDayEvent = all_day
            changed.append("all_day")
        if reminder_minutes is not None:
            apt.ReminderMinutesBeforeStart = reminder_minutes
            apt.ReminderSet = reminder_minutes > 0
            changed.append("reminder")
        if recipients is not None:
            apt.RequiredAttendees = recipients
            changed.append("recipients")

        if send_update and apt.MeetingStatus == 1:
            apt.Send()
        else:
            apt.Save()

        return {
            "message": "Appointment updated successfully.",
            "changed_fields": changed,
            "subject": apt.Subject,
        }

    def respond_to_invitation(
        self,
        entry_id: str,
        response: str,
        comment: str = "",
    ) -> dict[str, Any]:
        """Accept, decline, or tentatively accept a meeting invitation.

        Args:
            entry_id: The Outlook EntryID of the meeting request or appointment.
            response: 'accept', 'decline', or 'tentative'.
            comment: Optional message to include with the response.

        Returns:
            Dict with response result.
        """
        item = self.namespace.GetItemFromID(entry_id)

        # If it's a MeetingItem (class 53), get the associated appointment
        if item.Class == 53:  # olMeetingRequest
            item = item.GetAssociatedAppointment(False)

        response_lower = response.lower()
        # olMeetingAccepted=3, olMeetingDeclined=4, olMeetingTentative=5
        response_map = {"accept": 3, "decline": 4, "tentative": 5}
        if response_lower not in response_map:
            raise ValueError(
                f"Invalid response: '{response}'. Use 'accept', 'decline', or 'tentative'."
            )

        ol_response = response_map[response_lower]
        verb = {"accept": "Accepted", "decline": "Declined", "tentative": "Tentatively accepted"}[response_lower]

        try:
            item.Respond(ol_response, True)
        except Exception:
            # Fallback: set response state directly
            item.ResponseState = ol_response
            item.Save()

        return {
            "message": f"{verb} invitation: '{item.Subject}'.",
            "response": response_lower,
            "subject": item.Subject,
        }

    def get_free_busy(
        self,
        start_date: str | None = None,
        months: int = 1,
        account_name: str | None = None,
    ) -> dict[str, Any]:
        """Get free/busy information for the current user.

        Args:
            start_date: ISO date string (default: today).
            months: Number of months to query (default: 1).
            account_name: Optional account display name.

        Returns:
            Dict with free/busy data keyed by date.
        """
        if not start_date:
            start_date = datetime.date.today().isoformat()

        dt = datetime.datetime.fromisoformat(start_date)
        freebusy_str = ""

        # Try multiple approaches to get free/busy
        # Approach 1: CurrentUser.FreeBusy
        try:
            freebusy_str = self.namespace.CurrentUser.FreeBusy(
                dt, 30, True  # Start, MinPerChar, CompleteFormat
            )
        except Exception:
            pass

        # Approach 2: Create a recipient from current user
        if not freebusy_str:
            try:
                recipient = self.namespace.CreateRecipient(self.namespace.CurrentUser.Name)
                recipient.Resolve()
                if recipient.Resolved:
                    freebusy_str = recipient.FreeBusy(dt, 30, True)
            except Exception:
                pass

        # Approach 3: Try with first account's SMTP address
        if not freebusy_str:
            try:
                for acc in self.namespace.Accounts:
                    smtp = acc.SmtpAddress
                    if smtp:
                        recipient = self.namespace.CreateRecipient(smtp)
                        recipient.Resolve()
                        if recipient.Resolved:
                            freebusy_str = recipient.FreeBusy(dt, 30, True)
                            break
            except Exception:
                pass

        if not freebusy_str:
            return {"error": "Could not retrieve free/busy information. This feature may require an Exchange account.", "slots": []}

        status_map = {"0": "Free", "1": "Tentative", "2": "Busy", "3": "Out of Office", "4": "Working Elsewhere"}

        # Parse into date-timed slots
        slots = []
        for i, ch in enumerate(freebusy_str):
            slot_start = dt + datetime.timedelta(minutes=30 * i)
            slots.append({
                "time": slot_start.isoformat(),
                "status": status_map.get(ch, f"Unknown ({ch})"),
            })

        return {
            "start_date": start_date,
            "slot_minutes": 30,
            "total_slots": len(slots),
            "slots": slots,
        }

    # ── Contacts: Update / Export ────────────────────────────────────

    def update_contact(
        self,
        entry_id: str,
        full_name: str | None = None,
        email: str | None = None,
        phone: str | None = None,
        mobile: str | None = None,
        home_phone: str | None = None,
        company: str | None = None,
        job_title: str | None = None,
    ) -> dict[str, Any]:
        """Update an existing contact. Only provided fields are changed.

        Args:
            entry_id: The Outlook EntryID of the contact.
            full_name: New full name.
            email: New email address.
            phone: New business phone.
            mobile: New mobile phone.
            home_phone: New home phone.
            company: New company name.
            job_title: New job title.

        Returns:
            Dict with update result.
        """
        contact = self.namespace.GetItemFromID(entry_id)
        if contact.Class != 40:  # olContact
            raise ValueError(f"Item is not a contact (class={contact.Class}).")

        changed = []
        if full_name is not None:
            contact.FullName = full_name
            changed.append("full_name")
        if email is not None:
            contact.Email1Address = email
            changed.append("email")
        if phone is not None:
            contact.BusinessTelephoneNumber = phone
            changed.append("phone")
        if mobile is not None:
            contact.MobileTelephoneNumber = mobile
            changed.append("mobile")
        if home_phone is not None:
            contact.HomeTelephoneNumber = home_phone
            changed.append("home_phone")
        if company is not None:
            contact.CompanyName = company
            changed.append("company")
        if job_title is not None:
            contact.JobTitle = job_title
            changed.append("job_title")

        contact.Save()

        return {
            "message": "Contact updated successfully.",
            "changed_fields": changed,
            "full_name": contact.FullName,
        }

    def export_contacts(
        self,
        format: str = "csv",
        save_path: str = "",
        account_name: str | None = None,
    ) -> dict[str, Any]:
        """Export contacts to a file (CSV or vCard format).

        Args:
            format: 'csv' or 'vcard'.
            save_path: Directory to save the file (default: current directory).
            account_name: Optional account display name.

        Returns:
            Dict with export result.
        """
        contacts = self.list_contacts(account_name=account_name)

        if not save_path:
            save_path = os.getcwd()

        if format == "csv":
            filename = "outlook_contacts.csv"
            full_path = os.path.join(save_path, filename)
            if contacts:
                fieldnames = list(contacts[0].keys())
                with open(full_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.DictWriter(f, fieldnames=fieldnames)
                    writer.writeheader()
                    writer.writerows(contacts)
            else:
                with open(full_path, "w", newline="", encoding="utf-8-sig") as f:
                    f.write("No contacts found.\n")

        elif format == "vcard":
            filename = "outlook_contacts.vcf"
            full_path = os.path.join(save_path, filename)
            with open(full_path, "w", encoding="utf-8") as f:
                for c in contacts:
                    f.write("BEGIN:VCARD\r\nVERSION:3.0\r\n")
                    f.write(f"FN:{c['full_name']}\r\n")
                    f.write(f"N:{c['last_name']};{c['first_name']};;;\r\n")
                    if c["email"]:
                        f.write(f"EMAIL;TYPE=INTERNET:{c['email']}\r\n")
                    if c["email2"]:
                        f.write(f"EMAIL;TYPE=INTERNET:{c['email2']}\r\n")
                    if c["business_phone"]:
                        f.write(f"TEL;TYPE=WORK:{c['business_phone']}\r\n")
                    if c["mobile_phone"]:
                        f.write(f"TEL;TYPE=CELL:{c['mobile_phone']}\r\n")
                    if c["home_phone"]:
                        f.write(f"TEL;TYPE=HOME:{c['home_phone']}\r\n")
                    if c["company"]:
                        f.write(f"ORG:{c['company']}\r\n")
                    if c["job_title"]:
                        f.write(f"TITLE:{c['job_title']}\r\n")
                    f.write("END:VCARD\r\n")
        else:
            raise ValueError(f"Unsupported format: '{format}'. Use 'csv' or 'vcard'.")

        return {
            "message": f"Exported {len(contacts)} contacts to {full_path}.",
            "file": full_path,
            "count": len(contacts),
        }

    # ── Email: Flag / Categorize / Empty Deleted / Open ──────────────

    def flag_email(
        self,
        entry_id: str,
        flag: bool = True,
        due_date: str | None = None,
        reminder_date: str | None = None,
    ) -> dict[str, Any]:
        """Flag an email for follow-up, optionally with a due date and reminder.

        Args:
            entry_id: The Outlook EntryID of the email.
            flag: True to flag, False to remove flag.
            due_date: Optional ISO date string for the flag due date.
            reminder_date: Optional ISO datetime string for a reminder.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        if flag:
            item.FlagStatus = 2  # olFlagMarked
            if due_date:
                item.FlagDueBy = due_date
            if reminder_date:
                item.FlagRequest = "Follow up"
                item.ReminderSet = True
                item.ReminderTime = reminder_date
        else:
            # Clear the flag — only set FlagStatus, don't touch FlagDueBy
            item.FlagStatus = 0  # olNoFlag
            item.ReminderSet = False
            try:
                item.ClearTaskFlag()
            except Exception:
                pass
        try:
            item.Save()
        except Exception:
            # Item may have been modified; re-get and retry once
            item = self.namespace.GetItemFromID(entry_id)
            if flag:
                item.FlagStatus = 2
                if due_date:
                    item.FlagDueBy = due_date
            else:
                item.FlagStatus = 0
                item.ReminderSet = False
            item.Save()

        action = "Flagged" if flag else "Unflagged"
        return {"message": f"{action} email: '{item.Subject}'.", "subject": item.Subject}

    def categorize_email(
        self,
        entry_id: str,
        categories: str = "",
        action: str = "set",
    ) -> dict[str, Any]:
        """Add, remove, set, or clear categories on an email.

        Args:
            entry_id: The Outlook EntryID of the email.
            categories: Category name(s), semicolon-separated.
            action: 'set' (replace all), 'add' (append), 'remove' (remove specific), 'clear' (remove all).

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)

        def _apply_categories(mail):
            if action == "set":
                mail.Categories = categories
            elif action == "add":
                existing = mail.Categories or ""
                mail.Categories = existing + "; " + categories if existing else categories
            elif action == "remove":
                existing = mail.Categories or ""
                remove_set = {c.strip() for c in categories.split(";") if c.strip()}
                current = [c.strip() for c in existing.split(";") if c.strip()]
                mail.Categories = "; ".join(c for c in current if c not in remove_set)
            elif action == "clear":
                mail.Categories = ""
            else:
                raise ValueError(f"Invalid action: '{action}'. Use 'set', 'add', 'remove', or 'clear'.")

        _apply_categories(item)
        try:
            item.Save()
        except Exception:
            # Item may have been modified; re-get and retry once
            item = self.namespace.GetItemFromID(entry_id)
            _apply_categories(item)
            item.Save()

        return {
            "message": f"Categories updated for email: '{item.Subject}'.",
            "categories": item.Categories or "",
            "subject": item.Subject,
        }

    def empty_deleted_folder(self, account_name: str | None = None) -> dict[str, Any]:
        """Empty the Deleted Items folder.

        Args:
            account_name: Optional account display name.

        Returns:
            Dict with result info.
        """
        folder = self._get_folder(self.OL_FOLDER_DELETED, account_name)
        count = folder.Items.Count
        for i in range(count, 0, -1):
            try:
                folder.Items.Item(i).Delete()
            except Exception:
                pass

        return {"message": f"Deleted Items folder emptied ({count} items).", "items_deleted": count}

    def open_email(self, entry_id: str) -> dict[str, Any]:
        """Open an email in a separate Outlook window for review.

        Args:
            entry_id: The Outlook EntryID of the email.

        Returns:
            Dict with result info.
        """
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 43:
            raise ValueError(f"Item is not an email (class={item.Class}).")
        item.Display()
        return {"message": f"Opened email: '{item.Subject}'.", "subject": item.Subject}

    # ── Draft Management ─────────────────────────────────────────────

    def update_draft(
        self,
        entry_id: str,
        subject: str | None = None,
        body: str | None = None,
        to: str | None = None,
        cc: str | None = None,
        bcc: str | None = None,
        html_body: bool = False,
        attachments: list[str] | None = None,
        importance: int | None = None,
    ) -> dict[str, Any]:
        """Update an existing draft email.

        Args:
            entry_id: The Outlook EntryID of the draft.
            subject: New subject.
            body: New body text.
            to: New recipients.
            cc: New CC recipients.
            bcc: New BCC recipients.
            html_body: If True, treat body as HTML.
            attachments: Additional file paths to attach.
            importance: Importance level (0=Low, 1=Normal, 2=High).

        Returns:
            Dict with update result.
        """
        draft = self.namespace.GetItemFromID(entry_id)
        if draft.Class != 43:
            raise ValueError(f"Item is not an email (class={draft.Class}).")

        changed = []
        if subject is not None:
            draft.Subject = subject
            changed.append("subject")
        if body is not None:
            if html_body:
                draft.HTMLBody = body
            else:
                draft.Body = body
            changed.append("body")
        if to is not None:
            draft.To = to
            changed.append("to")
        if cc is not None:
            draft.CC = cc
            changed.append("cc")
        if bcc is not None:
            draft.BCC = bcc
            changed.append("bcc")
        if importance is not None:
            draft.Importance = importance
            changed.append("importance")
        if attachments:
            for filepath in attachments:
                full_path = os.path.abspath(filepath)
                if os.path.exists(full_path):
                    draft.Attachments.Add(full_path)
                    changed.append(f"attachment: {os.path.basename(filepath)}")

        draft.Save()

        return {
            "message": "Draft updated successfully.",
            "entry_id": entry_id,
            "subject": draft.Subject,
            "changed_fields": changed,
        }

    def send_draft(self, entry_id: str) -> dict[str, Any]:
        """Send an existing draft email.

        Args:
            entry_id: The Outlook EntryID of the draft to send.

        Returns:
            Dict with send result.
        """
        draft = self.namespace.GetItemFromID(entry_id)
        if draft.Class != 43:
            raise ValueError(f"Item is not an email (class={draft.Class}).")
        subject = draft.Subject
        draft.Send()
        return {"message": f"Draft sent: '{subject}'.", "subject": subject}

    # ── Tasks ────────────────────────────────────────────────────────

    def list_tasks(
        self,
        count: int = 20,
        offset: int = 0,
        fields: str | None = None,
        include_completed: bool = False,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """List tasks from the Outlook Tasks folder.

        Args:
            count: Maximum tasks to return (default: 20).
            offset: Number of tasks to skip (for pagination).
            fields: Comma-separated field names to include. When omitted, returns a compact summary.
            include_completed: If True, include completed tasks.
            account_name: Optional account display name.

        Returns:
            List of compact task summary dicts (no body).
        """
        folder = self._get_folder(self.OL_FOLDER_TASKS, account_name)
        items = folder.Items
        items.Sort("[DueDate]")

        if not include_completed:
            try:
                items = items.Restrict("@SQL=\"urn:schemas:httpmail:complete\" = 0")
            except Exception:
                pass

        result = []
        start = 1 + offset
        end = min(start + count, items.Count + 1)
        for i in range(start, end):
            try:
                item = items.Item(i)
                if item.Class == 48:  # olTask
                    d = self._task_to_summary(item)
                    result.append(self._filter_fields(d, fields))
            except Exception:
                continue

        return result

    def get_task_by_id(self, entry_id: str) -> dict[str, Any]:
        """Get full task details by EntryID."""
        item = self.namespace.GetItemFromID(entry_id)
        if item.Class != 48:
            raise ValueError(f"Item is not a task (class={item.Class}).")
        return self._task_to_dict(item)

    def create_task(
        self,
        subject: str,
        body: str = "",
        due_date: str | None = None,
        start_date: str | None = None,
        importance: int = IMPORTANCE_NORMAL,
        reminder_minutes: int = 0,
    ) -> dict[str, Any]:
        """Create a new task.

        Args:
            subject: Task subject/title.
            body: Task body/notes.
            due_date: ISO date string for due date.
            start_date: ISO date string for start date.
            importance: 0=Low, 1=Normal, 2=High.
            reminder_minutes: Minutes before due to remind (0 for none).

        Returns:
            Dict with created task info.
        """
        task = self.app.CreateItem(3)  # 3 = olTaskItem
        task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = due_date
        if start_date:
            task.StartDate = start_date
        task.Importance = importance
        if reminder_minutes > 0:
            task.ReminderSet = True
            task.ReminderMinutesBeforeStart = reminder_minutes
        task.Save()

        return {
            "message": "Task created successfully.",
            "subject": subject,
            "entry_id": task.EntryID,
        }

    def update_task(
        self,
        entry_id: str,
        subject: str | None = None,
        body: str | None = None,
        due_date: str | None = None,
        start_date: str | None = None,
        status: int | None = None,
        importance: int | None = None,
        reminder_minutes: int | None = None,
        percent_complete: int | None = None,
    ) -> dict[str, Any]:
        """Update an existing task. Only provided fields are changed.

        Args:
            entry_id: The Outlook EntryID of the task.
            subject: New subject.
            body: New body/notes.
            due_date: New due date (ISO date string).
            start_date: New start date (ISO date string).
            status: 0=NotStarted, 1=InProgress, 2=Complete, 3=Waiting, 4=Deferred.
            importance: 0=Low, 1=Normal, 2=High.
            reminder_minutes: Minutes before due to remind (0 for none).
            percent_complete: Completion percentage (0-100).

        Returns:
            Dict with update result.
        """
        task = self.namespace.GetItemFromID(entry_id)
        if task.Class != 48:
            raise ValueError(f"Item is not a task (class={task.Class}).")

        changed = []
        if subject is not None:
            task.Subject = subject
            changed.append("subject")
        if body is not None:
            task.Body = body
            changed.append("body")
        if due_date is not None:
            task.DueDate = due_date
            changed.append("due_date")
        if start_date is not None:
            task.StartDate = start_date
            changed.append("start_date")
        if status is not None:
            task.Status = status
            changed.append("status")
        if importance is not None:
            task.Importance = importance
            changed.append("importance")
        if reminder_minutes is not None:
            task.ReminderSet = reminder_minutes > 0
            if reminder_minutes > 0:
                task.ReminderMinutesBeforeStart = reminder_minutes
            changed.append("reminder")
        if percent_complete is not None:
            task.PercentComplete = percent_complete
            changed.append("percent_complete")

        task.Save()

        return {
            "message": "Task updated successfully.",
            "changed_fields": changed,
            "entry_id": entry_id,
        }

    def delete_task(self, entry_id: str) -> dict[str, Any]:
        """Delete a task by its EntryID.

        Args:
            entry_id: The Outlook EntryID of the task.

        Returns:
            Dict with result info.
        """
        task = self.namespace.GetItemFromID(entry_id)
        if task.Class != 48:
            raise ValueError(f"Item is not a task (class={task.Class}).")
        subject = task.Subject
        task.Delete()
        return {"message": f"Deleted task: '{subject}'."}

    def mark_task_complete(self, entry_id: str, complete: bool = True) -> dict[str, Any]:
        """Mark a task as complete or not started.

        Args:
            entry_id: The Outlook EntryID of the task.
            complete: True to mark complete, False to mark not started.

        Returns:
            Dict with result info.
        """
        task = self.namespace.GetItemFromID(entry_id)
        if task.Class != 48:
            raise ValueError(f"Item is not a task (class={task.Class}).")

        if complete:
            task.Status = 2  # olTaskComplete
            task.PercentComplete = 100
            # DateCompleted is auto-set by Outlook; don't set it directly
        else:
            task.Status = 0  # olTaskNotStarted
            task.PercentComplete = 0

        task.Save()

        verb = "completed" if complete else "reopened"
        return {"message": f"Task '{task.Subject}' marked as {verb}."}

    # ── Rules ────────────────────────────────────────────────────────

    def get_rules(self) -> dict[str, Any]:
        """Get all Outlook inbox rules.

        Returns:
            Dict with list of rules.
        """
        rules = []
        try:
            for rule in self.namespace.DefaultStore.GetRules():
                rules.append({
                    "name": rule.Name,
                    "enabled": rule.Enabled,
                    "execution_order": rule.ExecutionOrder,
                    "rule_type": "Receive" if rule.RuleType == 0 else "Send",
                })
        except Exception as e:
            return {"error": f"Could not retrieve rules: {e}", "rules": []}

        return {"rules": rules, "count": len(rules)}

    def create_rule(
        self,
        name: str,
        condition_type: str = "sender",
        condition_value: str = "",
        action_type: str = "move",
        action_value: str = "",
        enabled: bool = True,
        account_name: str | None = None,
    ) -> dict[str, Any]:
        """Create a simple inbox rule.

        Args:
            name: Rule name.
            condition_type: 'sender' (filter by sender) or 'subject' (filter by subject text).
            condition_value: Value for the condition (email address or subject keyword).
            action_type: 'move', 'mark_read', 'delete', or 'categorize'.
            action_value: Target folder name (for 'move') or category name (for 'categorize').
            enabled: Whether the rule is enabled.
            account_name: Optional account display name.

        Returns:
            Dict with creation result.
        """
        store = self.namespace.DefaultStore
        if account_name:
            for acc in self.namespace.Accounts:
                if acc.DisplayName.lower() == account_name.lower():
                    store = acc.DeliveryStore
                    break

        rules = store.GetRules()
        rule = rules.Create(name, 0)  # 0 = olRuleReceive

        # Set condition
        if condition_type == "sender":
            rule.Conditions.SenderAddress.Enabled = True
            rule.Conditions.SenderAddress.Address = [condition_value]
        elif condition_type == "subject":
            rule.Conditions.Subject.Enabled = True
            rule.Conditions.Subject.Text = [condition_value]
        else:
            raise ValueError(f"Invalid condition_type: '{condition_type}'. Use 'sender' or 'subject'.")

        # Set action
        if action_type == "move":
            # Find the target folder
            target_folder = None
            for folder in self.namespace.Folders:
                try:
                    target_folder = folder.Folders(action_value)
                    break
                except Exception:
                    # Try top-level default folders
                    for ftype_name, ftype_const in [
                        ("Inbox", self.OL_FOLDER_INBOX),
                        ("Sent", self.OL_FOLDER_SENT),
                        ("Drafts", self.OL_FOLDER_DRAFTS),
                        ("Deleted", self.OL_FOLDER_DELETED),
                        ("Calendar", self.OL_FOLDER_CALENDAR),
                        ("Contacts", self.OL_FOLDER_CONTACTS),
                        ("Tasks", self.OL_FOLDER_TASKS),
                    ]:
                        if action_value.lower() == ftype_name.lower():
                            target_folder = self._get_folder(ftype_const)
                            break
                    if target_folder:
                        break
            # Try recursing into subfolders of Inbox
            if not target_folder:
                inbox = self._get_folder(self.OL_FOLDER_INBOX)
                try:
                    target_folder = inbox.Folders(action_value)
                except Exception:
                    pass
            if target_folder:
                rule.Actions.MoveToFolder.Enabled = True
                rule.Actions.MoveToFolder.Folder = target_folder
            else:
                rules.Remove(rule.Name)
                rules.Save()
                raise ValueError(
                    f"Could not find folder '{action_value}'. "
                    f"Check the folder name or use a default folder name."
                )
        elif action_type == "mark_read":
            # MarkAsRead is read-only on RuleActions; get the action object, then enable it
            try:
                mark_action = rule.Actions.MarkAsRead
                mark_action.Enabled = True
            except Exception:
                # Fallback: use Item() by constant (olRuleActionMarkAsRead = 3)
                try:
                    rule.Actions.Item(3).Enabled = True
                except Exception:
                    # Final fallback: try direct property assignment
                    rule.Actions.MarkAsRead = True
        elif action_type == "delete":
            rule.Actions.Delete.Enabled = True
        elif action_type == "categorize":
            rule.Actions.AssignToCategory.Enabled = True
            rule.Actions.AssignToCategory.Categories = [action_value]
        else:
            rules.Remove(rule.Name)
            rules.Save()
            raise ValueError(
                f"Invalid action_type: '{action_type}'. Use 'move', 'mark_read', 'delete', or 'categorize'."
            )

        rule.Enabled = enabled
        rules.Save()

        return {
            "message": f"Rule '{name}' created successfully.",
            "name": name,
            "enabled": enabled,
        }

    # ── Helpers ──────────────────────────────────────────────────────

    @staticmethod
    def _strip_falsy(d: dict[str, Any]) -> dict[str, Any]:
        """Remove empty strings and None values to shrink serialized payload."""
        return {k: v for k, v in d.items() if v != "" and v is not None}

    @staticmethod
    def _filter_fields(d: dict[str, Any], fields: str | None) -> dict[str, Any]:
        """Filter dict to only include requested fields (comma-separated).

        If *fields* is None or empty, return the dict unchanged.
        """
        if not fields:
            return d
        allowed = set(f.strip() for f in fields.split(","))
        return {k: v for k, v in d.items() if k in allowed}

    def _mail_to_summary(self, item: Any) -> dict[str, Any]:
        """Convert a MailItem to a compact summary dict (no body, no attachment list)."""
        result: dict[str, Any] = {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "sender_name": item.SenderName,
            "received_time": str(item.ReceivedTime),
            "unread": item.UnRead,
            "importance": item.Importance,
            "has_attachments": item.Attachments.Count > 0,
            "attachment_count": item.Attachments.Count,
            "categories": item.Categories or "",
            "flag_status": item.FlagStatus,
        }
        return self._strip_falsy(result)

    def _mail_to_dict(self, item: Any, include_body: bool = False) -> dict[str, Any]:
        """Convert a MailItem COM object to a dict."""
        result: dict[str, Any] = {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "sender_name": item.SenderName,
            "sender_email": "",
            "to": item.To or "",
            "cc": item.CC or "",
            "received_time": str(item.ReceivedTime),
            "sent_on": str(item.SentOn) if item.SentOn else "",
            "unread": item.UnRead,
            "importance": item.Importance,
            "has_attachments": item.Attachments.Count > 0,
            "attachment_count": item.Attachments.Count,
            "size": item.Size,
            "categories": item.Categories or "",
            "flag_status": item.FlagStatus,
            "flag_due_date": str(item.FlagDueBy) if item.FlagDueBy else "",
        }

        # Try to get sender email
        try:
            if item.SenderEmailType == "SMTP":
                result["sender_email"] = item.SenderEmailAddress
            elif item.SenderEmailType == "EX":
                sender = item.Sender
                if sender:
                    try:
                        result["sender_email"] = sender.AddressEntryUserType
                    except Exception:
                        pass
        except Exception:
            pass

        if include_body:
            try:
                result["body"] = item.Body[:5000]  # Limit body length
            except Exception:
                result["body"] = ""

        if item.Attachments.Count > 0:
            attachments = []
            for i in range(1, item.Attachments.Count + 1):
                try:
                    att = item.Attachments.Item(i)
                    attachments.append({
                        "index": i,
                        "filename": att.FileName,
                        "size": att.Size,
                    })
                except Exception:
                    pass
            result["attachments"] = attachments

        return result

    def _appointment_to_summary(self, item: Any) -> dict[str, Any]:
        """Convert an AppointmentItem to a compact summary dict (no body)."""
        result: dict[str, Any] = {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "duration_minutes": item.Duration,
            "location": item.Location or "",
            "all_day": item.AllDayEvent,
            "organizer": item.Organizer or "",
        }
        return self._strip_falsy(result)

    def _appointment_to_dict(self, item: Any) -> dict[str, Any]:
        """Convert an AppointmentItem COM object to a dict."""
        response_map = {0: "None", 1: "Organized", 2: "Tentative", 3: "Accepted", 4: "Declined", 5: "NotResponded"}
        meeting_status_map = {0: "NonMeeting", 1: "Meeting", 2: "Received", 3: "Canceled"}
        return {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "start": str(item.Start),
            "end": str(item.End),
            "duration_minutes": item.Duration,
            "location": item.Location or "",
            "all_day": item.AllDayEvent,
            "body": item.Body[:500] if item.Body else "",
            "organizer": item.Organizer or "",
            "required_attendees": item.RequiredAttendees or "",
            "response_status": response_map.get(item.ResponseStatus, f"Unknown ({item.ResponseStatus})"),
            "meeting_status": meeting_status_map.get(item.MeetingStatus, f"Unknown ({item.MeetingStatus})"),
        }

    def _contact_to_summary(self, item: Any) -> dict[str, Any]:
        """Convert a ContactItem to a compact summary dict."""
        result: dict[str, Any] = {
            "entry_id": item.EntryID,
            "full_name": item.FullName or "",
            "email": item.Email1Address or "",
            "company": item.CompanyName or "",
            "job_title": item.JobTitle or "",
        }
        return self._strip_falsy(result)

    def _contact_to_dict(self, item: Any) -> dict[str, Any]:
        """Convert a ContactItem COM object to a dict."""
        return {
            "entry_id": item.EntryID,
            "full_name": item.FullName or "",
            "first_name": item.FirstName or "",
            "last_name": item.LastName or "",
            "email": item.Email1Address or "",
            "email2": item.Email2Address or "",
            "business_phone": item.BusinessTelephoneNumber or "",
            "mobile_phone": item.MobileTelephoneNumber or "",
            "home_phone": item.HomeTelephoneNumber or "",
            "company": item.CompanyName or "",
            "job_title": item.JobTitle or "",
        }

    def _task_to_summary(self, item: Any) -> dict[str, Any]:
        """Convert a TaskItem to a compact summary dict (no body)."""
        status_map = {0: "Not Started", 1: "In Progress", 2: "Complete", 3: "Waiting", 4: "Deferred"}
        result: dict[str, Any] = {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "due_date": str(item.DueDate) if item.DueDate else "",
            "start_date": str(item.StartDate) if item.StartDate else "",
            "status": status_map.get(item.Status, f"Unknown ({item.Status})"),
            "importance": item.Importance,
            "percent_complete": item.PercentComplete,
            "categories": item.Categories or "",
        }
        return self._strip_falsy(result)

    def _task_to_dict(self, item: Any) -> dict[str, Any]:
        """Convert a TaskItem COM object to a dict."""
        status_map = {0: "Not Started", 1: "In Progress", 2: "Complete", 3: "Waiting", 4: "Deferred"}
        return {
            "entry_id": item.EntryID,
            "subject": item.Subject,
            "body": item.Body[:1000] if item.Body else "",
            "due_date": str(item.DueDate) if item.DueDate else "",
            "start_date": str(item.StartDate) if item.StartDate else "",
            "date_completed": str(item.DateCompleted) if item.DateCompleted else "",
            "status": status_map.get(item.Status, f"Unknown ({item.Status})"),
            "status_code": item.Status,
            "importance": item.Importance,
            "percent_complete": item.PercentComplete,
            "total_work": item.TotalWork,
            "actual_work": item.ActualWork,
            "owner": item.Owner or "",
            "categories": item.Categories or "",
        }
