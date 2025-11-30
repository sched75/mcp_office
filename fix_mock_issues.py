"""Script to fix mock issues in Outlook tests."""

import os
import tempfile

def create_test_files():
    """Create temporary test files for attachment tests."""
    test_files = {
        "test_document.pdf": b"PDF test content",
        "test_report.pdf": b"Report test content",
        "test_file.txt": b"Text file content"
    }
    
    created_files = {}
    for filename, content in test_files.items():
        temp_file = tempfile.NamedTemporaryFile(
            suffix=f"_{filename}", 
            delete=False,
            dir=os.getcwd()
        )
        temp_file.write(content)
        temp_file.close()
        created_files[filename] = temp_file.name
    
    return created_files

def fix_mock_attachments():
    """Fix MockAttachments to have initial attachments."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockAttachments to have initial attachments
    old_attachments = """class MockAttachments:
    \"\"\"Mock attachments collection.\"\"\"

    def __init__(self):
        self.Count = 0
        self._attachments = []"""
    
    new_attachments = """class MockAttachments:
    \"\"\"Mock attachments collection.\"\"\"

    def __init__(self):
        self.Count = 1
        self._attachments = [MockAttachment("test_file.pdf")]"""
    
    content = content.replace(old_attachments, new_attachments)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_mock_items():
    """Fix MockItems to have Count attribute."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockItems Count attribute
    old_items = """class MockItems:
    \"\"\"Mock items collection.\"\"\"

    def __init__(self):
        self.Count = 10
        self._items = [MockMailItem() for _ in range(3)]"""
    
    new_items = """class MockItems:
    \"\"\"Mock items collection.\"\"\"

    def __init__(self):
        self.Count = 10
        self._items = [MockMailItem() for _ in range(3)]"""
    
    content = content.replace(old_items, new_items)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_mock_folder():
    """Fix MockFolder to have UnReadItemCount."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockFolder to have UnReadItemCount
    old_folder = """class MockFolder:
    \"\"\"Mock Outlook folder.\"\"\"

    def __init__(self, name):
        self.Name = name
        self.Items = MockItems()
        self.UnReadItemCount = 5
        self.Folders = MockFolders()"""
    
    new_folder = """class MockFolder:
    \"\"\"Mock Outlook folder.\"\"\"

    def __init__(self, name):
        self.Name = name
        self.Items = MockItems()
        self.UnReadItemCount = 5
        self.Folders = MockFolders()"""
    
    content = content.replace(old_folder, new_folder)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_mock_appointment():
    """Fix MockAppointmentItem to have Send method."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Add Send method to MockAppointmentItem
    old_appointment = """    def SaveAs(self, path, format_type):
        \"\"\"Save as file.\"\"\"
        pass

    def Respond(self, response_type, send_response):
        \"\"\"Respond to meeting.\"\"\"
        pass"""
    
    new_appointment = """    def SaveAs(self, path, format_type):
        \"\"\"Save as file.\"\"\"
        pass

    def Send(self):
        \"\"\"Send meeting.\"\"\"
        pass

    def Respond(self, response_type, send_response):
        \"\"\"Respond to meeting.\"\"\"
        pass"""
    
    content = content.replace(old_appointment, new_appointment)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_mock_contact():
    """Fix MockContactItem to have FullName attribute."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockContactItem to have FullName
    old_contact = """class MockContactItem(MockItem):
    \"\"\"Mock contact item.\"\"\"

    def __init__(self):
        super().__init__()
        self.FirstName = \"John\"
        self.LastName = \"Doe\"
        self.FullName = \"John Doe\"
        self.Email1Address = \"john.doe@example.com\"
        self.BusinessTelephoneNumber = \"+1234567890\"
        self.CompanyName = \"Test Company\"
        self.JobTitle = \"Test Title\""""
    
    new_contact = """class MockContactItem(MockItem):
    \"\"\"Mock contact item.\"\"\"

    def __init__(self):
        super().__init__()
        self.FirstName = \"John\"
        self.LastName = \"Doe\"
        self.FullName = \"John Doe\"
        self.Email1Address = \"john.doe@example.com\"
        self.BusinessTelephoneNumber = \"+1234567890\"
        self.CompanyName = \"Test Company\"
        self.JobTitle = \"Test Title\""""
    
    content = content.replace(old_contact, new_contact)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_mock_session():
    """Fix MockSession to have CreateRecipient method."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Add CreateRecipient to MockSession
    old_session = """class MockSession:
    \"\"\"Mock Outlook session.\"\"\"

    def __init__(self):
        self.Accounts = MockAccounts()"""
    
    new_session = """class MockSession:
    \"\"\"Mock Outlook session.\"\"\"

    def __init__(self):
        self.Accounts = MockAccounts()

    def CreateRecipient(self, email):
        \"\"\"Create recipient.\"\"\"
        return MockRecipient(email)"""
    
    content = content.replace(old_session, new_session)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_calendar_items():
    """Fix calendar items to use MockAppointmentItem instead of MockMailItem."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockItems to use MockAppointmentItem for calendar tests
    old_items_init = """        self._items = [MockMailItem() for _ in range(3)]"""
    new_items_init = """        self._items = [MockMailItem() for _ in range(2)] + [MockAppointmentItem()]"""
    
    content = content.replace(old_items_init, new_items_init)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_task_items():
    """Fix task items to use MockTaskItem instead of MockMailItem."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockItems to use MockTaskItem for task tests
    old_items_init = """        self._items = [MockMailItem() for _ in range(2)] + [MockAppointmentItem()]"""
    new_items_init = """        self._items = [MockMailItem() for _ in range(1)] + [MockAppointmentItem()] + [MockTaskItem()]"""
    
    content = content.replace(old_items_init, new_items_init)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def fix_contact_items():
    """Fix contact items to use MockContactItem instead of MockMailItem."""
    with open("tests/test_outlook_service.py", "r", encoding="utf-8") as f:
        content = f.read()
    
    # Fix MockItems to use MockContactItem for contact tests
    old_items_init = """        self._items = [MockMailItem() for _ in range(1)] + [MockAppointmentItem()] + [MockTaskItem()]"""
    new_items_init = """        self._items = [MockMailItem() for _ in range(1)] + [MockAppointmentItem()] + [MockTaskItem()] + [MockContactItem()]"""
    
    content = content.replace(old_items_init, new_items_init)
    
    with open("tests/test_outlook_service.py", "w", encoding="utf-8") as f:
        f.write(content)

def main():
    """Apply all fixes."""
    print("Creating test files...")
    test_files = create_test_files()
    
    print("Fixing mock attachments...")
    fix_mock_attachments()
    
    print("Fixing mock items...")
    fix_mock_items()
    
    print("Fixing mock folder...")
    fix_mock_folder()
    
    print("Fixing mock appointment...")
    fix_mock_appointment()
    
    print("Fixing mock contact...")
    fix_mock_contact()
    
    print("Fixing mock session...")
    fix_mock_session()
    
    print("Fixing calendar items...")
    fix_calendar_items()
    
    print("Fixing task items...")
    fix_task_items()
    
    print("Fixing contact items...")
    fix_contact_items()
    
    print("All fixes applied!")
    print(f"Created test files: {list(test_files.keys())}")
    
    # Write test file paths to a file for cleanup
    with open("test_files.txt", "w") as f:
        for filename, path in test_files.items():
            f.write(f"{filename}: {path}\n")

if __name__ == "__main__":
    main()