"""Attachment operations mixin for Outlook service.

This module provides attachment-related functionality (5 methods).
"""

from typing import Any

from ..core.exceptions import AttachmentError, OutlookItemNotFoundError
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from ..utils.validators import validate_file_path, validate_string_not_empty


class AttachmentOperationsMixin:
    """Mixin providing attachment operations for Outlook.

    Provides 5 methods for managing attachments:
    - add_attachment
    - list_attachments
    - save_attachment
    - remove_attachment
    - send_with_attachments
    """

    @com_safe("add_attachment")
    def add_attachment(
        self,
        email_entry_id: str,
        file_path: str,
        display_name: str | None = None,
    ) -> dict[str, Any]:
        """Add an attachment to an email.

        Args:
            email_entry_id: Entry ID of the email
            file_path: Path to the file to attach
            display_name: Display name for attachment (optional)

        Returns:
            Dictionary with result

        Raises:
            OutlookItemNotFoundError: If email not found
            AttachmentError: If attachment fails

        Example:
            >>> result = outlook.add_attachment(
            ...     email_entry_id="...",
            ...     file_path="C:/documents/report.pdf"
            ... )
        """
        path = validate_file_path(file_path, must_exist=True)

        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        try:
            attachment = mail_item.Attachments.Add(str(path))
            if display_name:
                attachment.DisplayName = display_name

            mail_item.Save()

            return dict_to_result(
                success=True,
                message="Attachment added successfully",
                file_path=str(path),
                display_name=display_name or path.name,
            )
        except Exception as e:
            raise AttachmentError("add", str(path), str(e)) from e

    @com_safe("list_attachments")
    def list_attachments(self, email_entry_id: str) -> dict[str, Any]:
        """List all attachments of an email.

        Args:
            email_entry_id: Entry ID of the email

        Returns:
            Dictionary with attachment list

        Raises:
            OutlookItemNotFoundError: If email not found

        Example:
            >>> result = outlook.list_attachments(email_entry_id="...")
            >>> for att in result['attachments']:
            ...     print(att['filename'])
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        attachments = []
        for i in range(1, mail_item.Attachments.Count + 1):
            att = mail_item.Attachments.Item(i)
            attachments.append(
                {
                    "index": i,
                    "filename": att.FileName,
                    "display_name": att.DisplayName,
                    "size": att.Size,
                    "type": att.Type,  # 1=File, 2=OLE, 3=Link, etc.
                }
            )

        return dict_to_result(
            success=True,
            message=f"Found {len(attachments)} attachment(s)",
            attachments=attachments,
            count=len(attachments),
        )

    @com_safe("save_attachment")
    def save_attachment(
        self,
        email_entry_id: str,
        attachment_index: int,
        save_path: str,
    ) -> dict[str, Any]:
        """Save an attachment from an email to disk.

        Args:
            email_entry_id: Entry ID of the email
            attachment_index: Index of the attachment (1-based)
            save_path: Path where to save the file

        Returns:
            Dictionary with result

        Raises:
            OutlookItemNotFoundError: If email not found
            AttachmentError: If save fails

        Example:
            >>> result = outlook.save_attachment(
            ...     email_entry_id="...",
            ...     attachment_index=1,
            ...     save_path="C:/downloads/file.pdf"
            ... )
        """
        validate_string_not_empty(save_path, "save_path")

        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        if attachment_index < 1 or attachment_index > mail_item.Attachments.Count:
            raise AttachmentError(
                "save",
                f"index {attachment_index}",
                f"Index out of range (1-{mail_item.Attachments.Count})",
            )

        try:
            attachment = mail_item.Attachments.Item(attachment_index)
            attachment.SaveAsFile(save_path)

            return dict_to_result(
                success=True,
                message="Attachment saved successfully",
                filename=attachment.FileName,
                save_path=save_path,
            )
        except Exception as e:
            raise AttachmentError("save", save_path, str(e)) from e

    @com_safe("remove_attachment")
    def remove_attachment(
        self,
        email_entry_id: str,
        attachment_index: int,
    ) -> dict[str, Any]:
        """Remove an attachment from an email.

        Args:
            email_entry_id: Entry ID of the email
            attachment_index: Index of the attachment (1-based)

        Returns:
            Dictionary with result

        Raises:
            OutlookItemNotFoundError: If email not found
            AttachmentError: If removal fails

        Example:
            >>> result = outlook.remove_attachment(
            ...     email_entry_id="...",
            ...     attachment_index=1
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        mail_item = namespace.GetItemFromID(email_entry_id)

        if mail_item is None:
            raise OutlookItemNotFoundError("email", email_entry_id)

        if attachment_index < 1 or attachment_index > mail_item.Attachments.Count:
            raise AttachmentError(
                "remove",
                f"index {attachment_index}",
                f"Index out of range (1-{mail_item.Attachments.Count})",
            )

        try:
            attachment = mail_item.Attachments.Item(attachment_index)
            filename = attachment.FileName
            attachment.Delete()
            mail_item.Save()

            return dict_to_result(
                success=True,
                message="Attachment removed successfully",
                filename=filename,
            )
        except Exception as e:
            raise AttachmentError("remove", f"index {attachment_index}", str(e)) from e

    @com_safe("send_with_attachments")
    def send_with_attachments(
        self,
        to: str | list[str],
        subject: str,
        body: str,
        attachments: list[str],
        cc: str | list[str] | None = None,
        bcc: str | list[str] | None = None,
        importance: int = 1,
    ) -> dict[str, Any]:
        """Send an email with multiple attachments.

        Args:
            to: Recipient email address(es)
            subject: Email subject
            body: Email body content
            attachments: List of file paths to attach
            cc: CC recipient(s) (optional)
            bcc: BCC recipient(s) (optional)
            importance: Email importance (0=Low, 1=Normal, 2=High)

        Returns:
            Dictionary with send result

        Raises:
            AttachmentError: If any attachment fails

        Example:
            >>> result = outlook.send_with_attachments(
            ...     to="recipient@example.com",
            ...     subject="Documents",
            ...     body="Please find attached documents",
            ...     attachments=["file1.pdf", "file2.docx"]
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

        # Add attachments
        attached_files = []
        for file_path in attachments:
            try:
                path = validate_file_path(file_path, must_exist=True)
                mail_item.Attachments.Add(str(path))
                attached_files.append(path.name)
            except Exception as e:
                raise AttachmentError("attach", file_path, str(e)) from e

        mail_item.Send()

        return dict_to_result(
            success=True,
            message="Email sent with attachments successfully",
            subject=subject,
            recipients=mail_item.To,
            attachments=attached_files,
            attachment_count=len(attached_files),
        )
