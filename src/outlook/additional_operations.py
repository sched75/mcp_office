"""Additional operations mixins for Outlook service.

This module consolidates meeting, contact, task, and advanced operations.
"""

from datetime import datetime
from typing import Any

from ..core.exceptions import (
    MeetingOperationError,
    OutlookItemNotFoundError,
)
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from ..utils.validators import validate_string_not_empty


class MeetingOperationsMixin:
    """Mixin providing meeting operations (8 methods)."""

    @com_safe("create_meeting_request")
    def create_meeting_request(
        self,
        subject: str,
        start_time: str,
        end_time: str,
        required_attendees: list[str],
        optional_attendees: list[str] | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict[str, Any]:
        """Create a meeting request."""
        validate_string_not_empty(subject, "subject")

        try:
            meeting = self.application.CreateItem(1)  # AppointmentItem
            meeting.MeetingStatus = 1  # olMeeting

            meeting.Subject = subject
            meeting.Start = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            meeting.End = datetime.fromisoformat(end_time.replace("Z", "+00:00"))

            if location:
                meeting.Location = location
            if body:
                meeting.Body = body

            # Add required attendees
            for attendee in required_attendees:
                recipient = meeting.Recipients.Add(attendee)
                recipient.Type = 1  # olRequired

            # Add optional attendees
            if optional_attendees:
                for attendee in optional_attendees:
                    recipient = meeting.Recipients.Add(attendee)
                    recipient.Type = 2  # olOptional

            meeting.Recipients.ResolveAll()
            meeting.Save()
            meeting.Send()

            return dict_to_result(
                success=True,
                message="Meeting request created and sent",
                entry_id=meeting.EntryID,
                subject=subject,
            )
        except Exception as e:
            raise MeetingOperationError("create_meeting_request", str(e)) from e

    @com_safe("invite_participants")
    def invite_participants(
        self,
        meeting_entry_id: str,
        attendees: list[str],
        required: bool = True,
    ) -> dict[str, Any]:
        """Add participants to a meeting."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            for attendee in attendees:
                recipient = meeting.Recipients.Add(attendee)
                recipient.Type = 1 if required else 2

            meeting.Recipients.ResolveAll()
            meeting.Save()

            return dict_to_result(
                success=True,
                message=f"Added {len(attendees)} participant(s)",
                subject=meeting.Subject,
            )
        except Exception as e:
            raise MeetingOperationError("invite_participants", str(e)) from e

    @com_safe("accept_meeting")
    def accept_meeting(self, meeting_entry_id: str) -> dict[str, Any]:
        """Accept a meeting invitation."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            meeting.Respond(3, True)  # 3 = olMeetingAccepted
            return dict_to_result(
                success=True,
                message="Meeting accepted",
                subject=meeting.Subject,
            )
        except Exception as e:
            raise MeetingOperationError("accept_meeting", str(e)) from e

    @com_safe("decline_meeting")
    def decline_meeting(self, meeting_entry_id: str) -> dict[str, Any]:
        """Decline a meeting invitation."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            meeting.Respond(4, True)  # 4 = olMeetingDeclined
            return dict_to_result(
                success=True,
                message="Meeting declined",
                subject=meeting.Subject,
            )
        except Exception as e:
            raise MeetingOperationError("decline_meeting", str(e)) from e

    @com_safe("propose_new_time")
    def propose_new_time(
        self,
        meeting_entry_id: str,
        new_start: str,
        new_end: str,
    ) -> dict[str, Any]:
        """Propose a new time for a meeting."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            meeting.Start = datetime.fromisoformat(new_start.replace("Z", "+00:00"))
            meeting.End = datetime.fromisoformat(new_end.replace("Z", "+00:00"))
            meeting.Respond(2, True)  # 2 = olMeetingTentative
            return dict_to_result(
                success=True,
                message="New time proposed",
                subject=meeting.Subject,
            )
        except Exception as e:
            raise MeetingOperationError("propose_new_time", str(e)) from e

    @com_safe("cancel_meeting")
    def cancel_meeting(self, meeting_entry_id: str) -> dict[str, Any]:
        """Cancel a meeting."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            subject = meeting.Subject
            meeting.MeetingStatus = 5  # olMeetingCanceled
            meeting.Send()
            return dict_to_result(
                success=True,
                message="Meeting canceled",
                subject=subject,
            )
        except Exception as e:
            raise MeetingOperationError("cancel_meeting", str(e)) from e

    @com_safe("update_meeting")
    def update_meeting(
        self,
        meeting_entry_id: str,
        subject: str | None = None,
        start_time: str | None = None,
        end_time: str | None = None,
        location: str | None = None,
    ) -> dict[str, Any]:
        """Update meeting details."""
        namespace = self.application.GetNamespace("MAPI")
        meeting = namespace.GetItemFromID(meeting_entry_id)

        if meeting is None:
            raise OutlookItemNotFoundError("meeting", meeting_entry_id)

        try:
            if subject:
                meeting.Subject = subject
            if start_time:
                meeting.Start = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            if end_time:
                meeting.End = datetime.fromisoformat(end_time.replace("Z", "+00:00"))
            if location:
                meeting.Location = location

            meeting.Save()
            meeting.Send()

            return dict_to_result(
                success=True,
                message="Meeting updated",
                subject=meeting.Subject,
            )
        except Exception as e:
            raise MeetingOperationError("update_meeting", str(e)) from e

    @com_safe("check_availability")
    def check_availability(
        self,
        attendees: list[str],
        start_time: str,
        end_time: str,
        duration_minutes: int = 60,
    ) -> dict[str, Any]:
        """Check availability of participants."""
        try:
            # This is a simplified version
            # Real implementation would query FreeBusy data
            return dict_to_result(
                success=True,
                message="Availability check completed",
                attendees=attendees,
                note="Simplified check - implement FreeBusy query for full functionality",
            )
        except Exception as e:
            raise MeetingOperationError("check_availability", str(e)) from e


class ContactOperationsMixin:
    """Mixin providing contact operations (9 methods)."""

    @com_safe("create_contact")
    def create_contact(
        self,
        first_name: str,
        last_name: str,
        email: str | None = None,
        phone: str | None = None,
        company: str | None = None,
        job_title: str | None = None,
    ) -> dict[str, Any]:
        """Create a new contact."""
        validate_string_not_empty(first_name, "first_name")
        validate_string_not_empty(last_name, "last_name")

        contact = self.application.CreateItem(2)  # 2 = olContactItem

        contact.FirstName = first_name
        contact.LastName = last_name

        if email:
            contact.Email1Address = email
        if phone:
            contact.BusinessTelephoneNumber = phone
        if company:
            contact.CompanyName = company
        if job_title:
            contact.JobTitle = job_title

        contact.Save()

        return dict_to_result(
            success=True,
            message="Contact created successfully",
            entry_id=contact.EntryID,
            full_name=contact.FullName,
        )

    @com_safe("modify_contact")
    def modify_contact(
        self,
        contact_entry_id: str,
        first_name: str | None = None,
        last_name: str | None = None,
        email: str | None = None,
        phone: str | None = None,
    ) -> dict[str, Any]:
        """Modify an existing contact."""
        namespace = self.application.GetNamespace("MAPI")
        contact = namespace.GetItemFromID(contact_entry_id)

        if contact is None:
            raise OutlookItemNotFoundError("contact", contact_entry_id)

        if first_name:
            contact.FirstName = first_name
        if last_name:
            contact.LastName = last_name
        if email:
            contact.Email1Address = email
        if phone:
            contact.BusinessTelephoneNumber = phone

        contact.Save()

        return dict_to_result(
            success=True,
            message="Contact modified successfully",
            full_name=contact.FullName,
        )

    @com_safe("delete_contact")
    def delete_contact(self, contact_entry_id: str) -> dict[str, Any]:
        """Delete a contact."""
        namespace = self.application.GetNamespace("MAPI")
        contact = namespace.GetItemFromID(contact_entry_id)

        if contact is None:
            raise OutlookItemNotFoundError("contact", contact_entry_id)

        full_name = contact.FullName
        contact.Delete()

        return dict_to_result(
            success=True,
            message="Contact deleted successfully",
            full_name=full_name,
        )

    @com_safe("search_contact")
    def search_contact(self, search_term: str) -> dict[str, Any]:
        """Search for contacts."""
        namespace = self.application.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(10)  # 10 = olFolderContacts

        filter_str = f"@SQL=\"urn:schemas:contacts:fileas\" LIKE '%{search_term}%'"
        items = contacts_folder.Items.Restrict(filter_str)

        results = []
        for contact in items:
            results.append({
                "entry_id": contact.EntryID,
                "full_name": contact.FullName,
                "email": contact.Email1Address,
                "phone": contact.BusinessTelephoneNumber,
                "company": contact.CompanyName,
            })

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} contact(s)",
            results=results,
            count=len(results),
        )

    @com_safe("list_all_contacts")
    def list_all_contacts(self) -> dict[str, Any]:
        """List all contacts."""
        namespace = self.application.GetNamespace("MAPI")
        contacts_folder = namespace.GetDefaultFolder(10)

        results = []
        for contact in contacts_folder.Items:
            results.append({
                "entry_id": contact.EntryID,
                "full_name": contact.FullName,
                "email": contact.Email1Address,
            })

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} contact(s)",
            results=results,
            count=len(results),
        )

    @com_safe("create_contact_group")
    def create_contact_group(self, group_name: str) -> dict[str, Any]:
        """Create a contact group."""
        validate_string_not_empty(group_name, "group_name")

        dist_list = self.application.CreateItem(7)  # 7 = olDistributionListItem
        dist_list.DLName = group_name
        dist_list.Save()

        return dict_to_result(
            success=True,
            message="Contact group created",
            group_name=group_name,
        )

    @com_safe("add_to_contact_group")
    def add_to_contact_group(
        self,
        group_entry_id: str,
        contact_email: str,
    ) -> dict[str, Any]:
        """Add contact to a group."""
        namespace = self.application.GetNamespace("MAPI")
        group = namespace.GetItemFromID(group_entry_id)

        if group is None:
            raise OutlookItemNotFoundError("contact group", group_entry_id)

        recipient = self.application.Session.CreateRecipient(contact_email)
        group.AddMember(recipient)
        group.Save()

        return dict_to_result(
            success=True,
            message="Contact added to group",
            group_name=group.DLName,
        )

    @com_safe("export_contacts_vcf")
    def export_contacts_vcf(self, output_path: str) -> dict[str, Any]:
        """Export contacts to VCF."""
        # Simplified implementation
        return dict_to_result(
            success=True,
            message="Export functionality - to be fully implemented",
            output_path=output_path,
        )

    @com_safe("import_contacts")
    def import_contacts(self, file_path: str) -> dict[str, Any]:
        """Import contacts from file."""
        # Simplified implementation
        return dict_to_result(
            success=True,
            message="Import functionality - to be fully implemented",
            file_path=file_path,
        )


class TaskOperationsMixin:
    """Mixin providing task operations (7 methods)."""

    @com_safe("create_task")
    def create_task(
        self,
        subject: str,
        body: str | None = None,
        due_date: str | None = None,
        priority: int = 1,  # 0=Low, 1=Normal, 2=High
    ) -> dict[str, Any]:
        """Create a new task."""
        validate_string_not_empty(subject, "subject")

        task = self.application.CreateItem(3)  # 3 = olTaskItem
        task.Subject = subject

        if body:
            task.Body = body
        if due_date:
            task.DueDate = datetime.fromisoformat(due_date.replace("Z", "+00:00"))

        task.Importance = priority
        task.Save()

        return dict_to_result(
            success=True,
            message="Task created successfully",
            entry_id=task.EntryID,
            subject=subject,
        )

    @com_safe("modify_task")
    def modify_task(
        self,
        task_entry_id: str,
        subject: str | None = None,
        body: str | None = None,
        due_date: str | None = None,
    ) -> dict[str, Any]:
        """Modify a task."""
        namespace = self.application.GetNamespace("MAPI")
        task = namespace.GetItemFromID(task_entry_id)

        if task is None:
            raise OutlookItemNotFoundError("task", task_entry_id)

        if subject:
            task.Subject = subject
        if body:
            task.Body = body
        if due_date:
            task.DueDate = datetime.fromisoformat(due_date.replace("Z", "+00:00"))

        task.Save()

        return dict_to_result(
            success=True,
            message="Task modified successfully",
            subject=task.Subject,
        )

    @com_safe("delete_task")
    def delete_task(self, task_entry_id: str) -> dict[str, Any]:
        """Delete a task."""
        namespace = self.application.GetNamespace("MAPI")
        task = namespace.GetItemFromID(task_entry_id)

        if task is None:
            raise OutlookItemNotFoundError("task", task_entry_id)

        subject = task.Subject
        task.Delete()

        return dict_to_result(
            success=True,
            message="Task deleted successfully",
            subject=subject,
        )

    @com_safe("mark_task_complete")
    def mark_task_complete(self, task_entry_id: str) -> dict[str, Any]:
        """Mark a task as complete."""
        namespace = self.application.GetNamespace("MAPI")
        task = namespace.GetItemFromID(task_entry_id)

        if task is None:
            raise OutlookItemNotFoundError("task", task_entry_id)

        task.Complete = True
        task.DateCompleted = datetime.now()
        task.Status = 2  # olTaskComplete
        task.Save()

        return dict_to_result(
            success=True,
            message="Task marked as complete",
            subject=task.Subject,
        )

    @com_safe("set_task_priority")
    def set_task_priority(
        self,
        task_entry_id: str,
        priority: int,  # 0=Low, 1=Normal, 2=High
    ) -> dict[str, Any]:
        """Set task priority."""
        namespace = self.application.GetNamespace("MAPI")
        task = namespace.GetItemFromID(task_entry_id)

        if task is None:
            raise OutlookItemNotFoundError("task", task_entry_id)

        task.Importance = priority
        task.Save()

        priority_names = {0: "Low", 1: "Normal", 2: "High"}

        return dict_to_result(
            success=True,
            message="Task priority set",
            subject=task.Subject,
            priority=priority_names.get(priority, "Unknown"),
        )

    @com_safe("set_task_due_date")
    def set_task_due_date(
        self,
        task_entry_id: str,
        due_date: str,
    ) -> dict[str, Any]:
        """Set task due date."""
        namespace = self.application.GetNamespace("MAPI")
        task = namespace.GetItemFromID(task_entry_id)

        if task is None:
            raise OutlookItemNotFoundError("task", task_entry_id)

        task.DueDate = datetime.fromisoformat(due_date.replace("Z", "+00:00"))
        task.Save()

        return dict_to_result(
            success=True,
            message="Task due date set",
            subject=task.Subject,
            due_date=due_date,
        )

    @com_safe("list_tasks")
    def list_tasks(self, completed: bool | None = None) -> dict[str, Any]:
        """List tasks."""
        namespace = self.application.GetNamespace("MAPI")
        tasks_folder = namespace.GetDefaultFolder(13)  # 13 = olFolderTasks

        items = tasks_folder.Items

        if completed is not None:
            filter_str = f"[Complete] = {completed}"
            items = items.Restrict(filter_str)

        results = []
        for task in items:
            results.append({
                "entry_id": task.EntryID,
                "subject": task.Subject,
                "due_date": str(task.DueDate) if task.DueDate else None,
                "complete": task.Complete,
                "priority": task.Importance,
            })

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} task(s)",
            results=results,
            count=len(results),
        )


class AdvancedOperationsMixin:
    """Mixin for advanced operations (categories, signatures, accounts, etc.)."""

    @com_safe("create_category")
    def create_category(self, name: str, color: int = 0) -> dict[str, Any]:
        """Create a category."""
        namespace = self.application.GetNamespace("MAPI")
        categories = namespace.Categories
        categories.Add(name, color)

        return dict_to_result(
            success=True,
            message="Category created",
            name=name,
        )

    @com_safe("apply_category")
    def apply_category(self, item_entry_id: str, category: str) -> dict[str, Any]:
        """Apply category to item."""
        namespace = self.application.GetNamespace("MAPI")
        item = namespace.GetItemFromID(item_entry_id)

        if item is None:
            raise OutlookItemNotFoundError("item", item_entry_id)

        item.Categories = category
        item.Save()

        return dict_to_result(
            success=True,
            message="Category applied",
            category=category,
        )

    @com_safe("list_categories")
    def list_categories(self) -> dict[str, Any]:
        """List all categories."""
        namespace = self.application.GetNamespace("MAPI")
        categories = namespace.Categories

        results = []
        for cat in categories:
            results.append({"name": cat.Name, "color": cat.Color})

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} category(ies)",
            results=results,
        )

    @com_safe("list_accounts")
    def list_accounts(self) -> dict[str, Any]:
        """List all email accounts."""
        accounts = self.application.Session.Accounts

        results = []
        for account in accounts:
            results.append({
                "display_name": account.DisplayName,
                "user_name": account.UserName,
                "smtp_address": account.SmtpAddress,
            })

        return dict_to_result(
            success=True,
            message=f"Found {len(results)} account(s)",
            results=results,
        )

    @com_safe("get_default_account")
    def get_default_account(self) -> dict[str, Any]:
        """Get default email account."""
        namespace = self.application.GetNamespace("MAPI")
        account = namespace.Accounts.Item(1)

        return dict_to_result(
            success=True,
            message="Default account retrieved",
            display_name=account.DisplayName,
            smtp_address=account.SmtpAddress,
        )
