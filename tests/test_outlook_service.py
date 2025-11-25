"""Tests for Outlook service."""

import pytest

from src.core.exceptions import (
    AttachmentError,
    COMInitializationError,
    OutlookItemNotFoundError,
)
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
        if item_type == 0:  # Mail
            item = MockMailItem()
        elif item_type == 1:  # Appointment
            item = MockAppointmentItem()
        elif item_type == 2:  # Contact
            item = MockContactItem()
        elif item_type == 3:  # Task
            item = MockTaskItem()
        else:
            item = MockItem()

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
        self.UnReadItemCount = 5
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
        self._folders = [MockFolder("Subfolder1"), MockFolder("Subfolder2")]

    def __call__(self, name):
        """Get folder by name."""
        return MockFolder(name)

    def Add(self, name):
        """Add folder."""
        return MockFolder(name)

    def __iter__(self):
        """Iterate folders."""
        return iter(self._folders)


class MockItems:
    """Mock items collection."""

    def __init__(self):
        self.Count = 10
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
    """Mock Outlook item base."""

    def __init__(self):
        self.EntryID = "test_entry_id"
        self.Subject = "Test Subject"
        self.Body = "Test Body"
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
        self.To = ""
        self.CC = ""
        self.BCC = ""
        self.SenderName = "Test Sender"
        self.SenderEmailAddress = "sender@example.com"
        self.ReceivedTime = "2024-01-01T12:00:00"
        self.SentOn = "2024-01-01T12:00:00"
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
        self.Start = "2024-01-15T10:00:00"
        self.End = "2024-01-15T11:00:00"
        self.Location = "Test Location"
        self.Organizer = "organizer@example.com"
        self.RequiredAttendees = ""
        self.OptionalAttendees = ""
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
        self.FullName = "John Doe"
        self.Email1Address = "john.doe@example.com"
        self.BusinessTelephoneNumber = "+1234567890"
        self.CompanyName = "Test Company"
        self.JobTitle = "Test Title"


class MockTaskItem(MockItem):
    """Mock task item."""

    def __init__(self):
        super().__init__()
        self.DueDate = "2024-12-31T23:59:59"
        self.Importance = 1
        self.Complete = False
        self.DateCompleted = None
        self.Status = 0


class MockAttachments:
    """Mock attachments collection."""

    def __init__(self):
        self.Count = 0
        self._attachments = []

    def Add(self, file_path):
        """Add attachment."""
        att = MockAttachment(file_path)
        self._attachments.append(att)
        self.Count += 1
        return att

    def Item(self, index):
        """Get attachment by index."""
        return self._attachments[index - 1]


class MockAttachment:
    """Mock attachment."""

    def __init__(self, file_path):
        self.FileName = file_path.split("/")[-1].split("\\")[-1]
        self.DisplayName = self.FileName
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


class MockRecurrencePattern:
    """Mock recurrence pattern."""

    def __init__(self):
        self.RecurrenceType = 0
        self.Interval = 1
        self.Occurrences = 0
        self.PatternEndDate = None


@pytest.fixture
def mock_outlook_app(monkeypatch):
    """Mock Outlook COM object."""

    def mock_dispatch(app_name):
        if app_name == "Outlook.Application":
            return MockOutlookApp()
        raise Exception(f"Unknown application: {app_name}")

    import win32com.client

    monkeypatch.setattr(win32com.client, "Dispatch", mock_dispatch)


@pytest.fixture
def outlook_service(mock_outlook_app):
    """Create Outlook service instance."""
    service = OutlookService()
    service.initialize()
    return service


class TestOutlookService:
    """Test suite for Outlook service."""

    def test_initialization(self, outlook_service):
        """Test service initialization."""
        assert outlook_service.is_initialized
        assert outlook_service.application is not None

    def test_create_document(self, outlook_service):
        """Test create document (initialize)."""
        result = outlook_service.create_document()
        assert result["success"]
        assert "accounts" in result

    def test_send_email(self, outlook_service):
        """Test sending email."""
        result = outlook_service.send_email(
            to="recipient@example.com",
            subject="Test Email",
            body="This is a test email",
        )

        assert result["success"]
        assert result["subject"] == "Test Email"

    def test_create_appointment(self, outlook_service):
        """Test creating appointment."""
        result = outlook_service.create_appointment(
            subject="Team Meeting",
            start_time="2024-01-15T10:00:00",
            end_time="2024-01-15T11:00:00",
            location="Conference Room A",
        )

        assert result["success"]
        assert result["subject"] == "Team Meeting"

    def test_create_contact(self, outlook_service):
        """Test creating contact."""
        result = outlook_service.create_contact(
            first_name="John",
            last_name="Doe",
            email="john.doe@example.com",
        )

        assert result["success"]
        assert "entry_id" in result

    def test_create_task(self, outlook_service):
        """Test creating task."""
        result = outlook_service.create_task(
            subject="Complete project",
            body="Finish the project by end of month",
        )

        assert result["success"]
        assert result["subject"] == "Complete project"

    def test_list_folders(self, outlook_service):
        """Test listing folders."""
        result = outlook_service.list_folders("Inbox", recursive=False)

        assert result["success"]
        assert "folders" in result
        assert result["count"] >= 0

    def test_get_inbox_count(self, outlook_service):
        """Test getting inbox count."""
        result = outlook_service.get_inbox_count()

        assert result["success"]
        assert "total_items" in result
        assert "unread_items" in result

    def test_list_accounts(self, outlook_service):
        """Test listing accounts."""
        result = outlook_service.list_accounts()

        assert result["success"]
        assert "results" in result

    def test_create_category(self, outlook_service):
        """Test creating category."""
        result = outlook_service.create_category("Work", color=0)

        assert result["success"]
        assert result["name"] == "Work"

    def test_item_not_found_error(self, outlook_service):
        """Test item not found error."""
        with pytest.raises(OutlookItemNotFoundError):
            outlook_service.read_email("invalid")


class TestMailOperations:
    """Test mail operations."""

    def test_reply_to_email(self, outlook_service):
        """Test replying to email."""
        # Create a mock email first
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = mail_item.EntryID

        result = outlook_service.reply_to_email(
            email_entry_id=entry_id,
            body="Thank you for your email",
        )

        assert result["success"]

    def test_mark_as_read(self, outlook_service):
        """Test marking email as read."""
        mail_item = outlook_service.application.CreateItem(0)
        entry_id = mail_item.EntryID

        result = outlook_service.mark_as_read(entry_id)

        assert result["success"]


class TestCalendarOperations:
    """Test calendar operations."""

    def test_create_recurring_event(self, outlook_service):
        """Test creating recurring event."""
        result = outlook_service.create_recurring_event(
            subject="Weekly Meeting",
            start_time="2024-01-15T10:00:00",
            end_time="2024-01-15T11:00:00",
            recurrence_type=1,  # Weekly
            interval=1,
            occurrences=10,
        )

        assert result["success"]
        assert "entry_id" in result


class TestFolderOperations:
    """Test folder operations."""

    def test_create_folder(self, outlook_service):
        """Test creating folder."""
        result = outlook_service.create_folder(
            folder_name="Projects",
            parent_folder="Inbox",
        )

        assert result["success"]
        assert result["folder_name"] == "Projects"
