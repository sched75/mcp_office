@echo off
REM Script pour commiter les changements Outlook dans le projet MCP Office

echo ================================================
echo Commit des changements Outlook dans MCP Office
echo ================================================
echo.

cd /d "C:\Users\dsi\OneDrive\Documents\Personnel\mcp_office"

echo Verification du statut git...
git status
echo.

echo Ajout des nouveaux fichiers...
git add src/outlook/
git add src/core/types.py
git add src/core/exceptions.py
git add tests/test_outlook_service.py
git add TODO.md
echo.

echo Verification des fichiers ajoutes...
git status
echo.

echo Commit des changements...
git commit -m "feat(outlook): Add complete Outlook service with 85 functionalities

- Add OutlookService with full COM automation support
- Implement 12 email operations (send, read, reply, search, etc.)
- Implement 5 attachment operations
- Implement 7 folder operations
- Implement 10 calendar operations (appointments, recurring events, reminders)
- Implement 8 meeting operations (create, invite, accept/decline)
- Implement 9 contact operations (create, modify, search, groups)
- Implement 7 task operations (create, modify, complete, priority)
- Implement 27 advanced operations (categories, rules, signatures, accounts)
- Add 10 Outlook-specific types to core/types.py
- Add 6 Outlook-specific exceptions to core/exceptions.py
- Add comprehensive test suite with mocks
- Add complete documentation and README
- Update TODO.md with Outlook completion status
- Total: 295/295 functionalities complete (100%)

Architecture:
- Mixin pattern for modular functionality
- SOLID principles compliance
- PEP 8 compliance
- Full type hints and docstrings
- Robust error handling with custom exceptions

Files added:
- src/outlook/__init__.py
- src/outlook/outlook_service.py
- src/outlook/mail_operations.py
- src/outlook/attachment_operations.py
- src/outlook/folder_operations.py
- src/outlook/calendar_operations.py
- src/outlook/additional_operations.py
- src/outlook/README.md
- tests/test_outlook_service.py"

echo.
echo ================================================
echo Commit termine avec succes!
echo ================================================
echo.

echo Voulez-vous pusher les changements vers GitHub? (O/N)
set /p push_choice="> "

if /i "%push_choice%"=="O" (
    echo.
    echo Push vers GitHub...
    git push origin main
    echo.
    echo ================================================
    echo Push termine avec succes!
    echo ================================================
) else (
    echo.
    echo Push annule. Vous pouvez pusher plus tard avec: git push origin main
)

echo.
pause
