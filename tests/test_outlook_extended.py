"""Tests supplémentaires pour améliorer la couverture Outlook."""

import pytest

from src.outlook.outlook_service import OutlookService


class MockOutlookApp:
    """Mock Outlook application for testing."""

    def __init__(self):
        self.Visible = False
        self.DisplayAlerts = False
        self.Session = MockSession()
        self._items_created = []

    def CreateItem(self, item_type):
        """Create a mock item."""
        if item_type == 0:
            item = MockMailItem()
        elif item_type == 1:
            item = MockAppointmentItem()
        elif item_type == 2:
            item = MockContactItem()
        elif item_type == 3:
            item = MockTaskItem()
        else:
            item = MockMailItem()

        self._items_created.append(item)
        return item

    def GetNamespace(self, name):
        """Get namespace."""
        return MockNamespace()

    def Quit(self):
        """Quit application."""
        pass


class MockSession:
    """Mock Outlook session."""

    def __init__(self):
        self.Accounts = MockAccounts()


class MockAccounts:
    """Mock accounts collection."""

    def __init__(self):
        self.Count = 2

    def Item(self, index):
        """Get account by index."""
        return MockAccount()


class MockAccount:
    """Mock Outlook account."""

    def __init__(self):
        self.DisplayName = "Test Account"
        self.UserName = "test@example.com"
        self.SmtpAddress = "test@example.com"


class MockNamespace:
    """Mock MAPI namespace."""

    def __init__(self):
        self.Categories = MockCategories()
        self.Accounts = MockAccounts()

    def GetDefaultFolder(self, folder_type):
        """Get default folder."""
        return MockFolder(f"Folder_{folder_type}")

    def GetItemFromID(self, entry_id):
        """Get item by ID."""
        if entry_id == "invalid":
            return None
        # Return appropriate item based on entry_id prefix
        if entry_id.startswith("mail"):
            return MockMailItem()
        elif entry_id.startswith("appt"):
            return MockAppointmentItem()
        elif entry_id.startswith("contact"):
            return MockContactItem()
        elif entry_id.startswith("task"):
            return MockTaskItem()
        return MockMailItem()


class MockCategories:
    """Mock categories collection."""

    def __init__(self):
        self._categories = []

    def Add(self, name, color):
        """Add category."""
        cat = MockCategory(name, color)
        self._categories.append(cat)
        return cat

    def __iter__(self):
        """Iterate categories."""
        return iter(self._categories)


class MockCategory:
    """Mock category."""

    def __init__(self, name, color):
        self.Name = name
        self.Color = color


class MockFolder:
    """Mock Outlook folder."""

    def __init__(self, name):
        self.Name = name
        self.Items = MockItems()
        self.Folders = MockFolders()

    def Add(self, folder_name):
        """Add subfolder."""
        return MockFolder(folder_name)

    def Delete(self):
        """Delete folder."""
        pass

    def MoveTo(self, target):
        """Move folder."""
        pass


class MockFolders:
    """Mock folders collection."""

    def __init__(self):
        self._folders = [MockFolder("SubFolder1"), MockFolder("SubFolder2")]

    def __call__(self, name):
        """Get folder by name."""
        return MockFolder(name)

    def Add(self, name):
        """Add folder."""
        return MockFolder(name)


class MockItems:
    """Mock items collection."""

    def __init__(self):
        self._items = [MockMailItem() for _ in range(3)]

    def Restrict(self, filter_str):
        """Apply filter."""
        return self

    def Sort(self, sort_field):
        """Sort items."""
        pass

    def __iter__(self):
        """Iterate items."""
        return iter(self._items)


class MockItem:
    """Base mock item."""

    def __init__(self):
        self.EntryID = "mock_entry_id"
        self.Categories = ""

    def Save(self):
        """Save item."""
        pass

    def Delete(self):
        """Delete item."""
        pass


class MockMailItem(MockItem):
    """Mock mail item."""

    def __init__(self):
        super().__init__()
        self.Subject = "Test Email"
        self.Body = "Test body content"
        self.To = "recipient@example.com"
        self.CC = ""
        self.BCC = ""
        self.SenderName = "Test Sender"
        self.SenderEmailAddress = "sender@example.com"
        self.ReceivedTime = "2024-01-15 10:00:00"
        self.SentOn = "2024-01-15 09:00:00"
        self.Importance = 1
        self.Sensitivity = 0
        self.UnRead = True
        self.FlagStatus = 0
        self.Attachments = MockAttachments()
        self.Recipients = MockRecipients()

    def Send(self):
        """Send email."""
        pass

    def Reply(self):
        """Reply to email."""
        return MockMailItem()

    def ReplyAll(self):
        """Reply all."""
        return MockMailItem()

    def Forward(self):
        """Forward email."""
        return MockMailItem()

    def Move(self, folder):
        """Move to folder."""
        pass


class MockAppointmentItem(MockItem):
    """Mock appointment item."""

    def __init__(self):
        super().__init__()
        self.Subject = "Test Meeting"
        self.Start = "2024-01-15 10:00:00"
        self.End = "2024-01-15 11:00:00"
        self.Location = "Conference Room"
        self.Body = "Meeting agenda"
        self.Organizer = "organizer@example.com"
        self.RequiredAttendees = "attendee1@example.com"
        self.OptionalAttendees = "attendee2@example.com"
        self.BusyStatus = 2
        self.ReminderSet = True
        self.ReminderMinutesBeforeStart = 15
        self.IsRecurring = False
        self.MeetingStatus = 0
        self.Recipients = MockRecipients()

    def GetRecurrencePattern(self):
        """Get recurrence pattern."""
        return MockRecurrencePattern()

    def SaveAs(self, path, format_type):
        """Save as file."""
        pass

    def Respond(self, response_type, send_response):
        """Respond to meeting."""
        pass


class MockContactItem(MockItem):
    """Mock contact item."""

    def __init__(self):
        super().__init__()
        self.FirstName = "John"
        self.LastName = "Doe"
        self.Email1Address = "john.doe@example.com"
        self.BusinessTelephoneNumber = "+1-555-0100"
        self.CompanyName = "Test Company"
        self.JobTitle = "Manager"


class MockTaskItem(MockItem):
    """Mock task item."""

    def __init__(self):
        super().__init__()
        self.Subject = "Test Task"
        self.Body = "Task description"
        self.DueDate = "2024-01-30 17:00:00"
        self.Complete = False
        self.Importance = 1  # Priority


class MockRecurrencePattern:
    """Mock recurrence pattern."""

    def __init__(self):
        self.RecurrenceType = 0
        self.Interval = 1
        self.Occurrences = 0
        self.PatternEndDate = None


class MockAttachments:
    """Mock attachments collection."""

    def __init__(self):
        self.Count = 2
        self._attachments = []

    def Add(self, file_path):
        """Add attachment."""
        att = MockAttachment(file_path)
        self._attachments.append(att)
        return att

    def Item(self, index):
        """Get attachment by index."""
        return self._attachments[index - 1]


class MockAttachment:
    """Mock attachment."""

    def __init__(self, filename):
        self.FileName = filename
        self.DisplayName = filename
        self.Size = 1024
        self.Type = 1

    def SaveAsFile(self, path):
        """Save attachment."""
        pass

    def Delete(self):
        """Delete attachment."""
        pass


class MockRecipients:
    """Mock recipients collection."""

    def __init__(self):
        self._recipients = []

    def Add(self, email):
        """Add recipient."""
        recipient = MockRecipient(email)
        self._recipients.append(recipient)
        return recipient

    def ResolveAll(self):
        """Resolve all recipients."""
        return True


class MockRecipient:
    """Mock recipient."""

    def __init__(self, email):
        self.Address = email
        self.Type = 1  # olTo


@pytest.fixture
def outlook_service(monkeypatch):
    """Create Outlook service with mocked application."""
    mock_app = MockOutlookApp()

    def mock_dispatch(app_name):
        if app_name == "Outlook.Application":
            return mock_app
        raise Exception(f"Unknown application: {app_name}")

    monkeypatch.setattr("win32com.client.Dispatch", mock_dispatch)

    service = OutlookService()
    service.initialize()
    service.application = mock_app

    yield service

    service.cleanup()


# ============================================================================
# MAIL OPERATIONS TESTS
# ============================================================================


class TestMailOperationsExtended:
    """Extended tests for mail operations."""

    def test_forward_email(self, outlook_service):
        """Test forwarding email."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.forward_email(
            email_entry_id=entry_id,
            to="colleague@example.com",
            body="FYI",
        )

        assert result["success"]

    def test_reply_all_to_email(self, outlook_service):
        """Test reply all to email."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.reply_all_to_email(
            email_entry_id=entry_id,
            body="Thanks everyone",
        )

        assert result["success"]

    def test_read_email(self, outlook_service):
        """Test reading email."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.read_email(entry_id)

        assert result["success"]
        assert "subject" in result
        assert "sender" in result

    def test_mark_as_unread(self, outlook_service):
        """Test marking email as unread."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.mark_as_unread(entry_id)

        assert result["success"]

    def test_flag_email(self, outlook_service):
        """Test flagging email."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.flag_email(entry_id, flag_status=2)

        assert result["success"]

    def test_delete_email(self, outlook_service):
        """Test deleting email."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.delete_email(entry_id)

        assert result["success"]

    def test_move_email_to_folder(self, outlook_service):
        """Test moving email to folder."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.move_email_to_folder(
            email_entry_id=entry_id,
            folder_path="Inbox/Archive",
        )

        assert result["success"]

    def test_search_emails(self, outlook_service):
        """Test searching emails."""
        result = outlook_service.search_emails(
            subject="project",
            sender="boss@example.com",
            unread_only=True,
            max_results=10,
        )

        assert result["success"]
        assert "results" in result


# ============================================================================
# ATTACHMENT OPERATIONS TESTS
# ============================================================================


class TestAttachmentOperations:
    """Tests for attachment operations."""

    def test_add_attachment(self, outlook_service):
        """Test adding attachment."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.add_attachment(
            email_entry_id=entry_id,
            file_path="C:\\test\\document.pdf",
        )

        assert result["success"]

    def test_list_attachments(self, outlook_service):
        """Test listing attachments."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.list_attachments(entry_id)

        assert result["success"]
        assert "attachments" in result

    def test_save_attachment(self, outlook_service):
        """Test saving attachment."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.save_attachment(
            email_entry_id=entry_id,
            attachment_index=1,
            save_path="C:\\test\\saved_file.pdf",
        )

        assert result["success"]

    def test_remove_attachment(self, outlook_service):
        """Test removing attachment."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.remove_attachment(
            email_entry_id=entry_id,
            attachment_index=1,
        )

        assert result["success"]

    def test_send_with_attachments(self, outlook_service):
        """Test sending email with attachments."""
        result = outlook_service.send_with_attachments(
            to="recipient@example.com",
            subject="Report",
            body="Please find attached",
            attachments=["C:\\test\\report.pdf", "C:\\test\\data.xlsx"],
        )

        assert result["success"]


# ============================================================================
# FOLDER OPERATIONS TESTS
# ============================================================================


class TestFolderOperationsExtended:
    """Extended tests for folder operations."""

    def test_delete_folder(self, outlook_service):
        """Test deleting folder."""
        result = outlook_service.delete_folder("Inbox/OldStuff")

        assert result["success"]

    def test_rename_folder(self, outlook_service):
        """Test renaming folder."""
        result = outlook_service.rename_folder(
            folder_path="Inbox/Projects",
            new_name="ActiveProjects",
        )

        assert result["success"]

    def test_move_folder(self, outlook_service):
        """Test moving folder."""
        result = outlook_service.move_folder(
            folder_path="Inbox/Projects",
            destination_path="Archive",
        )

        assert result["success"]

    def test_get_folder_item_count(self, outlook_service):
        """Test getting folder item count."""
        result = outlook_service.get_folder_item_count("Inbox")

        assert result["success"]
        assert "count" in result

    def test_get_unread_count(self, outlook_service):
        """Test getting unread count."""
        result = outlook_service.get_unread_count("Inbox")

        assert result["success"]
        assert "unread_count" in result


# ============================================================================
# CALENDAR OPERATIONS TESTS
# ============================================================================


class TestCalendarOperationsExtended:
    """Extended tests for calendar operations."""

    def test_modify_appointment(self, outlook_service):
        """Test modifying appointment."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.modify_appointment(
            appointment_entry_id=entry_id,
            subject="Updated Meeting",
            location="New Room",
        )

        assert result["success"]

    def test_delete_appointment(self, outlook_service):
        """Test deleting appointment."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.delete_appointment(entry_id)

        assert result["success"]

    def test_read_appointment(self, outlook_service):
        """Test reading appointment."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.read_appointment(entry_id)

        assert result["success"]
        assert "subject" in result

    def test_search_appointments(self, outlook_service):
        """Test searching appointments."""
        result = outlook_service.search_appointments(
            subject="meeting",
            location="Conference Room",
            max_results=20,
        )

        assert result["success"]
        assert "results" in result

    def test_get_appointments_by_date(self, outlook_service):
        """Test getting appointments by date."""
        result = outlook_service.get_appointments_by_date(
            start_date="2024-01-15T00:00:00",
            end_date="2024-01-15T23:59:59",
        )

        assert result["success"]

    def test_set_reminder(self, outlook_service):
        """Test setting reminder."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.set_reminder(entry_id, reminder_minutes=30)

        assert result["success"]

    def test_set_busy_status(self, outlook_service):
        """Test setting busy status."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.set_busy_status(entry_id, busy_status=2)

        assert result["success"]

    def test_export_appointment_ics(self, outlook_service):
        """Test exporting appointment to ICS."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.export_appointment_ics(
            appointment_entry_id=entry_id,
            output_path="C:\\test\\meeting.ics",
        )

        assert result["success"]


# ============================================================================
# MEETING OPERATIONS TESTS
# ============================================================================


class TestMeetingOperations:
    """Tests for meeting operations."""

    def test_invite_participants(self, outlook_service):
        """Test inviting participants."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.invite_participants(
            meeting_entry_id=entry_id,
            attendees=["colleague1@example.com", "colleague2@example.com"],
            required=True,
        )

        assert result["success"]

    def test_accept_meeting(self, outlook_service):
        """Test accepting meeting."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.accept_meeting(entry_id)

        assert result["success"]

    def test_decline_meeting(self, outlook_service):
        """Test declining meeting."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.decline_meeting(entry_id)

        assert result["success"]

    def test_propose_new_time(self, outlook_service):
        """Test proposing new time."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.propose_new_time(
            meeting_entry_id=entry_id,
            new_start="2024-01-16T10:00:00",
            new_end="2024-01-16T11:00:00",
        )

        assert result["success"]

    def test_cancel_meeting(self, outlook_service):
        """Test canceling meeting."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.cancel_meeting(entry_id)

        assert result["success"]

    def test_update_meeting(self, outlook_service):
        """Test updating meeting."""
        appt = outlook_service.application.CreateItem(1)
        entry_id = "appt_" + appt.EntryID

        result = outlook_service.update_meeting(
            meeting_entry_id=entry_id,
            subject="Updated Meeting Title",
        )

        assert result["success"]


# ============================================================================
# CONTACT OPERATIONS TESTS
# ============================================================================


class TestContactOperations:
    """Tests for contact operations."""

    def test_modify_contact(self, outlook_service):
        """Test modifying contact."""
        contact = outlook_service.application.CreateItem(2)
        entry_id = "contact_" + contact.EntryID

        result = outlook_service.modify_contact(
            contact_entry_id=entry_id,
            first_name="Jane",
            last_name="Smith",
        )

        assert result["success"]

    def test_delete_contact(self, outlook_service):
        """Test deleting contact."""
        contact = outlook_service.application.CreateItem(2)
        entry_id = "contact_" + contact.EntryID

        result = outlook_service.delete_contact(entry_id)

        assert result["success"]

    def test_search_contact(self, outlook_service):
        """Test searching contact."""
        result = outlook_service.search_contact("John Doe")

        assert result["success"]
        assert "results" in result

    def test_list_all_contacts(self, outlook_service):
        """Test listing all contacts."""
        result = outlook_service.list_all_contacts()

        assert result["success"]
        assert "results" in result

    def test_create_contact_group(self, outlook_service):
        """Test creating contact group."""
        result = outlook_service.create_contact_group("Team Members")

        assert result["success"]

    def test_add_to_contact_group(self, outlook_service):
        """Test adding to contact group."""
        group = outlook_service.application.CreateItem(7)  # DistributionListItem
        entry_id = "contact_" + group.EntryID

        result = outlook_service.add_to_contact_group(
            group_entry_id=entry_id,
            contact_email="member@example.com",
        )

        assert result["success"]


# ============================================================================
# TASK OPERATIONS TESTS
# ============================================================================


class TestTaskOperations:
    """Tests for task operations."""

    def test_modify_task(self, outlook_service):
        """Test modifying task."""
        task = outlook_service.application.CreateItem(3)
        entry_id = "task_" + task.EntryID

        result = outlook_service.modify_task(
            task_entry_id=entry_id,
            subject="Updated Task",
            body="New description",
        )

        assert result["success"]

    def test_delete_task(self, outlook_service):
        """Test deleting task."""
        task = outlook_service.application.CreateItem(3)
        entry_id = "task_" + task.EntryID

        result = outlook_service.delete_task(entry_id)

        assert result["success"]

    def test_mark_task_complete(self, outlook_service):
        """Test marking task complete."""
        task = outlook_service.application.CreateItem(3)
        entry_id = "task_" + task.EntryID

        result = outlook_service.mark_task_complete(entry_id)

        assert result["success"]

    def test_set_task_priority(self, outlook_service):
        """Test setting task priority."""
        task = outlook_service.application.CreateItem(3)
        entry_id = "task_" + task.EntryID

        result = outlook_service.set_task_priority(entry_id, priority=2)

        assert result["success"]

    def test_set_task_due_date(self, outlook_service):
        """Test setting task due date."""
        task = outlook_service.application.CreateItem(3)
        entry_id = "task_" + task.EntryID

        result = outlook_service.set_task_due_date(
            task_entry_id=entry_id,
            due_date="2024-02-01T17:00:00",
        )

        assert result["success"]

    def test_list_tasks(self, outlook_service):
        """Test listing tasks."""
        result = outlook_service.list_tasks(completed=False)

        assert result["success"]
        assert "results" in result


# ============================================================================
# ADVANCED OPERATIONS TESTS
# ============================================================================


class TestAdvancedOperations:
    """Tests for advanced operations."""

    def test_apply_category(self, outlook_service):
        """Test applying category."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = "mail_" + mail_item.EntryID

        result = outlook_service.apply_category(entry_id, "Work")

        assert result["success"]

    def test_list_categories(self, outlook_service):
        """Test listing categories."""
        result = outlook_service.list_categories()

        assert result["success"]

    def test_get_default_account(self, outlook_service):
        """Test getting default account."""
        result = outlook_service.get_default_account()

        assert result["success"]
        assert "display_name" in result
