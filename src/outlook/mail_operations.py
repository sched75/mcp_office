"""Email operations mixin for Outlook service.

This module provides all email-related functionality (12 methods).
"""

from typing import Any

from ..core.exceptions import InvalidRecipientError, OutlookItemNotFoundError
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from ..utils.validators import validate_string_not_empty


class MailOperationsMixin:
    """Mixin providing email operations for Outlook.

    Provides 12 methods for managing emails:
    - create_new_message
    - send_email
    - reply_to_email
    - reply_all_to_email
    - forward_email
    - read_email
    - mark_as_read
    - mark_as_unread
    - flag_email
    - delete_email
    - move_email_to_folder
    - search_emails
    """

    @com_safe("create_new_message")
    def create_new_message(self) -> dict[str, Any]:
        """Create a new email message.

        Returns:
            Dictionary with creation result and message object info

        Example:
            >>> result = outlook.create_new_message()
            >>> result['success']
            True
        """
        mail_item = self.application.CreateItem(0)  # 0 = olMailItem

        return dict_to_result(
            success=True,
            message="Email message created successfully",
            item_id=mail_item.EntryID if mail_item.EntryID else "unsaved",
        )

    @com_safe("send_email")
    def send_email(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        cc: str | list[str] | None = None,
        bcc: str | list[str] | None = None,
        importance: int = 1,  # 0=Low, 1=Normal, 2=High
    ) -> dict[str, Any]:
        """Send an email.

        Args:
            to: Recipient email address(es)
            subject: Email subject
            body: Email body content
            cc: CC recipient(s) (optional)
            bcc: BCC recipient(s) (optional)
            importance: Email importance (0=Low, 1=Normal, 2=High)

        Returns:
            Dictionary with send result

        Raises:
            InvalidRecipientError: If recipient email is invalid

        Example:
            >>> result = outlook.send_email(
            ...     to="recipient@example.com",
            ...     subject="Test Email",
            ...     body="This is a test email"
            ... )
        """
        validate_string_not_empty(subject, "subject")
        validate_string_not_empty(body, "body")

        mail_item = self.application.CreateItem(0)

        # Set recipients
        if isinstance(to, list):
            mail_item.To = "; ".join(to)
        else:
            mail_item.To = to

        if cc:
            mail_item.CC = "; ".join(cc) if isinstance(cc, list) else cc

        if bcc:
            mail_item.BCC = "; ".join(bcc) if isinstance(bcc, list) else bcc

        mail_item.Subject = subject
        mail_item.Body = body
        mail_item.Importance = importance

        mail_item.Send()

        return dict_to_result(
            success=True,
            message="Email sent successfully",
            subject=subject,
            recipients=mail_item.To,
        )

    @com_safe("reply_to_email")
    def reply_to_email(
        self,
        email_entry_id: str,
        body: str,
        send_immediately: bool = True,
    ) -> dict[str, Any]:
        """Reply to an email.

        Args:
            email_entry_id: Entry ID of the email to reply to
            body: Reply body content
            send_immediately: Whether to send immediately or save as draft

        Returns:
            Dictionary with reply result

        Raises:
            OutlookItemNotFoundError: If email not found

        Example:
            >>> result = outlook.reply_to_email(
            ...     email_entry_id="...",
            ...     body="Thank you for your email"
            ... )
        """
        validate_string_not_empty(body, "body")

        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        reply = mail_item.Reply()
        reply.Body = body + "\n\n" + reply.Body

        if send_immediately:
            reply.Send()
            action = "sent"
        else:
            reply.Save()
            action = "saved as draft"

        return dict_to_result(
            success=True,
            message=f"Reply {action} successfully",
            original_subject=mail_item.Subject,
        )

    @com_safe("reply_all_to_email")
    def reply_all_to_email(
        self,
        email_entry_id: str,
        body: str,
        send_immediately: bool = True,
    ) -> dict[str, Any]:
        """Reply to all recipients of an email.

        Args:
            email_entry_id: Entry ID of the email
            body: Reply body content
            send_immediately: Whether to send immediately

        Returns:
            Dictionary with reply result

        Example:
            >>> result = outlook.reply_all_to_email(
            ...     email_entry_id="...",
            ...     body="Thank you all"
            ... )
        """
        validate_string_not_empty(body, "body")

        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        reply = mail_item.ReplyAll()
        reply.Body = body + "\n\n" + reply.Body

        if send_immediately:
            reply.Send()
            action = "sent"
        else:
            reply.Save()
            action = "saved as draft"

        return dict_to_result(
            success=True,
            message=f"Reply all {action} successfully",
            original_subject=mail_item.Subject,
        )

    @com_safe("forward_email")
    def forward_email(
        self,
        email_entry_id: str,
        to: str | list[str],
        body: str | None = None,
        send_immediately: bool = True,
    ) -> dict[str, Any]:
        """Forward an email.

        Args:
            email_entry_id: Entry ID of the email to forward
            to: Recipient(s) to forward to
            body: Additional body text (optional)
            send_immediately: Whether to send immediately

        Returns:
            Dictionary with forward result

        Example:
            >>> result = outlook.forward_email(
            ...     email_entry_id="...",
            ...     to="colleague@example.com"
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        forward = mail_item.Forward()

        if isinstance(to, list):
            forward.To = "; ".join(to)
        else:
            forward.To = to

        if body:
            forward.Body = body + "\n\n" + forward.Body

        if send_immediately:
            forward.Send()
            action = "sent"
        else:
            forward.Save()
            action = "saved as draft"

        return dict_to_result(
            success=True,
            message=f"Email forwarded {action} successfully",
            original_subject=mail_item.Subject,
            recipients=forward.To,
        )

    @com_safe("read_email")
    def read_email(self, email_entry_id: str) -> dict[str, Any]:
        """Read an email and retrieve its properties.

        Args:
            email_entry_id: Entry ID of the email

        Returns:
            Dictionary with email properties

        Raises:
            OutlookItemNotFoundError: If email not found

        Example:
            >>> result = outlook.read_email(email_entry_id="...")
            >>> print(result['subject'])
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        return dict_to_result(
            success=True,
            message="Email read successfully",
            entry_id=mail_item.EntryID,
            subject=mail_item.Subject,
            sender=mail_item.SenderEmailAddress,
            sender_name=mail_item.SenderName,
            to=mail_item.To,
            cc=mail_item.CC,
            bcc=mail_item.BCC,
            body=mail_item.Body,
            html_body=mail_item.HTMLBody if hasattr(mail_item, "HTMLBody") else None,
            received_time=str(mail_item.ReceivedTime),
            sent_on=str(mail_item.SentOn) if mail_item.SentOn else None,
            importance=mail_item.Importance,
            sensitivity=mail_item.Sensitivity,
            unread=mail_item.UnRead,
            has_attachments=mail_item.Attachments.Count > 0,
            attachment_count=mail_item.Attachments.Count,
        )

    @com_safe("mark_as_read")
    def mark_as_read(self, email_entry_id: str) -> dict[str, Any]:
        """Mark an email as read.

        Args:
            email_entry_id: Entry ID of the email

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.mark_as_read(email_entry_id="...")
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        mail_item.UnRead = False
        mail_item.Save()

        return dict_to_result(
            success=True,
            message="Email marked as read",
            subject=mail_item.Subject,
        )

    @com_safe("mark_as_unread")
    def mark_as_unread(self, email_entry_id: str) -> dict[str, Any]:
        """Mark an email as unread.

        Args:
            email_entry_id: Entry ID of the email

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.mark_as_unread(email_entry_id="...")
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        mail_item.UnRead = True
        mail_item.Save()

        return dict_to_result(
            success=True,
            message="Email marked as unread",
            subject=mail_item.Subject,
        )

    @com_safe("flag_email")
    def flag_email(
        self,
        email_entry_id: str,
        flag_status: int = 2,  # 0=Clear, 1=Complete, 2=Flagged
    ) -> dict[str, Any]:
        """Flag an email.

        Args:
            email_entry_id: Entry ID of the email
            flag_status: Flag status (0=Clear, 1=Complete, 2=Flagged)

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.flag_email(email_entry_id="...", flag_status=2)
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        mail_item.FlagStatus = flag_status
        mail_item.Save()

        flag_names = {0: "cleared", 1: "complete", 2: "flagged"}

        return dict_to_result(
            success=True,
            message=f"Email {flag_names.get(flag_status, 'updated')} successfully",
            subject=mail_item.Subject,
        )

    @com_safe("delete_email")
    def delete_email(self, email_entry_id: str) -> dict[str, Any]:
        """Delete an email.

        Args:
            email_entry_id: Entry ID of the email

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.delete_email(email_entry_id="...")
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        subject = mail_item.Subject
        mail_item.Delete()

        return dict_to_result(
            success=True,
            message="Email deleted successfully",
            subject=subject,
        )

    @com_safe("move_email_to_folder")
    def move_email_to_folder(
        self,
        email_entry_id: str,
        folder_path: str,
    ) -> dict[str, Any]:
        """Move an email to a different folder.

        Args:
            email_entry_id: Entry ID of the email
            folder_path: Path to the destination folder

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.move_email_to_folder(
            ...     email_entry_id="...",
            ...     folder_path="Inbox/Projects"
            ... )
        """
        validate_string_not_empty(folder_path, "folder_path")

        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        # Navigate to destination folder
        folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
        for folder_name in folder_path.split("/"):
            folder = folder.Folders(folder_name)

        mail_item.Move(folder)

        return dict_to_result(
            success=True,
            message="Email moved successfully",
            subject=mail_item.Subject,
            destination=folder_path,
        )

    @com_safe("search_emails")
    def search_emails(
        self,
        folder_name: str = "Inbox",
        subject: str | None = None,
        sender: str | None = None,
        body_contains: str | None = None,
        start_date: str | None = None,
        end_date: str | None = None,
        unread_only: bool = False,
        max_results: int = 50,
    ) -> dict[str, Any]:
        """Search for emails with various criteria.

        Args:
            folder_name: Folder to search in
            subject: Subject text to search for
            sender: Sender email address to filter by
            body_contains: Text to search in email body
            start_date: Start date for search (ISO format)
            end_date: End date for search (ISO format)
            unread_only: Only return unread emails
            max_results: Maximum number of results

        Returns:
            Dictionary with search results

        Example:
            >>> result = outlook.search_emails(
            ...     subject="project",
            ...     sender="boss@example.com",
            ...     unread_only=True
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox

        if folder_name != "Inbox":
            folder = folder.Folders(folder_name)

        # Build filter string
        filters = []
        if subject:
            filters.append(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject}%'")
        if sender:
            filters.append(f"@SQL=\"urn:schemas:httpmail:fromname\" LIKE '%{sender}%'")
        if unread_only:
            filters.append("@SQL=\"urn:schemas:httpmail:read\" = 0")
        if start_date:
            filters.append(
                f"@SQL=\"urn:schemas:httpmail:datereceived\" >= '{start_date}'"
            )
        if end_date:
            filters.append(f"@SQL=\"urn:schemas:httpmail:datereceived\" <= '{end_date}'")

        filter_string = " AND ".join(filters) if filters else None

        items = folder.Items
        if filter_string:
            items = items.Restrict(filter_string)

        results = []
        count = 0
        for item in items:
            if count >= max_results:
                break

            # Additional body search if specified
            if body_contains and body_contains.lower() not in item.Body.lower():
                continue

            results.append({
                "entry_id": item.EntryID,
                "subject": item.Subject,
                "sender": item.SenderName,
                "sender_email": item.SenderEmailAddress,
                "received_time": str(item.ReceivedTime),
                "unread": item.UnRead,
                "has_attachments": item.Attachments.Count > 0,
            })
            count += 1

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} email(s)",
            results=results,
            count=len(results),
        )
