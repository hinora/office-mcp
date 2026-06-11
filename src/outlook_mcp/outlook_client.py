"""
Microsoft Outlook COM client.

Provides a high-level Python interface to automate Microsoft Outlook
using the COM automation API (Outlook.Application).
"""

from __future__ import annotations

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
        count: int = 50,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List recent emails from a folder.

        Args:
            folder_type: The Outlook folder constant (default: Inbox).
            count: Maximum number of emails to return.
            account_name: Optional account display name to target a specific account.

        Returns:
            List of email dicts with subject, sender, received time, etc.
        """
        folder = self._get_folder(folder_type, account_name)
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # Sort descending

        result = []
        for i in range(1, min(count + 1, items.Count + 1)):
            try:
                item = items.Item(i)
                # Only process MailItem (class 43)
                if item.Class == 43:
                    result.append(self._mail_to_dict(item))
            except Exception:
                continue

        return result

    def search_emails(
        self,
        query: str = "",
        folder_type: int = OL_FOLDER_INBOX,
        count: int = 50,
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
            count: Maximum results.
            subject: Filter by subject containing this text.
            sender: Filter by sender name/email containing this text.
            received_after: ISO date string (e.g., '2026-01-01').
            received_before: ISO date string.
            unread_only: Only return unread emails.
            account_name: Optional account display name.

        Returns:
            List of matching email dicts.
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
        for i in range(1, min(count + 1, filtered.Count + 1)):
            try:
                item = filtered.Item(i)
                if item.Class == 43:
                    result.append(self._mail_to_dict(item))
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
        count: int = 50,
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List calendar events.

        Args:
            start_date: ISO date string for range start (default: today).
            end_date: ISO date string for range end (default: +30 days).
            count: Maximum events to return.
            account_name: Optional account display name.

        Returns:
            List of event dicts.
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
        for i in range(1, min(count + 1, filtered.Count + 1)):
            try:
                item = filtered.Item(i)
                if item.Class == 26:  # olAppointment
                    result.append(self._appointment_to_dict(item))
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
        count: int = 100,
        search: str = "",
        account_name: str | None = None,
    ) -> list[dict[str, Any]]:
        """
        List contacts.

        Args:
            count: Maximum contacts to return.
            search: Optional text to filter contacts (searches name and email).
            account_name: Optional account display name.

        Returns:
            List of contact dicts.
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
        for i in range(1, min(count + 1, items.Count + 1)):
            try:
                item = items.Item(i)
                if item.Class == 40:  # olContact
                    result.append(self._contact_to_dict(item))
            except Exception:
                continue

        return result

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

    # ── Helpers ──────────────────────────────────────────────────────

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

    def _appointment_to_dict(self, item: Any) -> dict[str, Any]:
        """Convert an AppointmentItem COM object to a dict."""
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
        }

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
