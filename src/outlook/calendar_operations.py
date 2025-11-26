"""Calendar operations mixin for Outlook service.

This module provides calendar-related functionality (10 methods).
"""

from datetime import datetime
from typing import Any

from ..core.exceptions import CalendarOperationError, OutlookItemNotFoundError
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from ..utils.validators import validate_string_not_empty


class CalendarOperationsMixin:
    """Mixin providing calendar operations for Outlook.

    Provides 10 methods for managing calendar:
    - create_appointment
    - modify_appointment
    - delete_appointment
    - read_appointment
    - create_recurring_event
    - search_appointments
    - get_appointments_by_date
    - set_reminder
    - set_busy_status
    - export_appointment_ics
    """

    @com_safe("create_appointment")
    def create_appointment(
        self,
        subject: str,
        start_time: str,
        end_time: str,
        location: str | None = None,
        body: str | None = None,
        reminder_minutes: int = 15,
        busy_status: int = 2,  # 0=Free, 1=Tentative, 2=Busy, 3=OutOfOffice
    ) -> dict[str, Any]:
        """Create a new appointment.

        Args:
            subject: Appointment subject
            start_time: Start time (ISO format)
            end_time: End time (ISO format)
            location: Location of appointment
            body: Appointment description
            reminder_minutes: Reminder time in minutes before start
            busy_status: Busy status (0=Free, 1=Tentative, 2=Busy, 3=OutOfOffice)

        Returns:
            Dictionary with creation result

        Example:
            >>> result = outlook.create_appointment(
            ...     subject="Team Meeting",
            ...     start_time="2024-01-15T10:00:00",
            ...     end_time="2024-01-15T11:00:00",
            ...     location="Conference Room A"
            ... )
        """
        validate_string_not_empty(subject, "subject")
        validate_string_not_empty(start_time, "start_time")
        validate_string_not_empty(end_time, "end_time")

        try:
            appt = self.application.CreateItem(1)  # 1 = olAppointmentItem

            appt.Subject = subject
            appt.Start = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            appt.End = datetime.fromisoformat(end_time.replace("Z", "+00:00"))

            if location:
                appt.Location = location
            if body:
                appt.Body = body

            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
            appt.BusyStatus = busy_status

            appt.Save()

            return dict_to_result(
                success=True,
                message="Appointment created successfully",
                entry_id=appt.EntryID,
                subject=subject,
                start_time=start_time,
                end_time=end_time,
                location=location,
            )
        except Exception as e:
            raise CalendarOperationError("create_appointment", str(e)) from e

    @com_safe("modify_appointment")
    def modify_appointment(
        self,
        appointment_entry_id: str,
        subject: str | None = None,
        start_time: str | None = None,
        end_time: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict[str, Any]:
        """Modify an existing appointment.

        Args:
            appointment_entry_id: Entry ID of the appointment
            subject: New subject (optional)
            start_time: New start time (optional)
            end_time: New end time (optional)
            location: New location (optional)
            body: New body (optional)

        Returns:
            Dictionary with modification result

        Example:
            >>> result = outlook.modify_appointment(
            ...     appointment_entry_id="...",
            ...     location="Virtual Meeting"
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        try:
            if subject is not None:
                appt.Subject = subject
            if start_time is not None:
                appt.Start = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            if end_time is not None:
                appt.End = datetime.fromisoformat(end_time.replace("Z", "+00:00"))
            if location is not None:
                appt.Location = location
            if body is not None:
                appt.Body = body

            appt.Save()

            return dict_to_result(
                success=True,
                message="Appointment modified successfully",
                entry_id=appointment_entry_id,
                subject=appt.Subject,
            )
        except Exception as e:
            raise CalendarOperationError("modify_appointment", str(e)) from e

    @com_safe("delete_appointment")
    def delete_appointment(self, appointment_entry_id: str) -> dict[str, Any]:
        """Delete an appointment.

        Args:
            appointment_entry_id: Entry ID of the appointment

        Returns:
            Dictionary with deletion result

        Example:
            >>> result = outlook.delete_appointment(appointment_entry_id="...")
        """
        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        try:
            subject = appt.Subject
            appt.Delete()

            return dict_to_result(
                success=True,
                message="Appointment deleted successfully",
                subject=subject,
            )
        except Exception as e:
            raise CalendarOperationError("delete_appointment", str(e)) from e

    @com_safe("read_appointment")
    def read_appointment(self, appointment_entry_id: str) -> dict[str, Any]:
        """Read appointment details.

        Args:
            appointment_entry_id: Entry ID of the appointment

        Returns:
            Dictionary with appointment details

        Example:
            >>> result = outlook.read_appointment(appointment_entry_id="...")
            >>> print(result['subject'])
        """
        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        return dict_to_result(
            success=True,
            message="Appointment read successfully",
            entry_id=appt.EntryID,
            subject=appt.Subject,
            start_time=str(appt.Start),
            end_time=str(appt.End),
            location=appt.Location,
            body=appt.Body,
            organizer=appt.Organizer,
            required_attendees=appt.RequiredAttendees
            if hasattr(appt, "RequiredAttendees")
            else None,
            optional_attendees=appt.OptionalAttendees
            if hasattr(appt, "OptionalAttendees")
            else None,
            busy_status=appt.BusyStatus,
            reminder_set=appt.ReminderSet,
            reminder_minutes=appt.ReminderMinutesBeforeStart if appt.ReminderSet else 0,
            is_recurring=appt.IsRecurring,
        )

    @com_safe("create_recurring_event")
    def create_recurring_event(
        self,
        subject: str,
        start_time: str,
        end_time: str,
        recurrence_type: int,  # 0=Daily, 1=Weekly, 2=Monthly, 3=Yearly
        interval: int = 1,
        occurrences: int | None = None,
        end_date: str | None = None,
        location: str | None = None,
        body: str | None = None,
    ) -> dict[str, Any]:
        """Create a recurring appointment.

        Args:
            subject: Appointment subject
            start_time: Start time (ISO format)
            end_time: End time (ISO format)
            recurrence_type: Type of recurrence (0=Daily, 1=Weekly, 2=Monthly, 3=Yearly)
            interval: Interval between occurrences
            occurrences: Number of occurrences (optional)
            end_date: End date for recurrence (optional)
            location: Location of appointment
            body: Appointment description

        Returns:
            Dictionary with creation result

        Example:
            >>> result = outlook.create_recurring_event(
            ...     subject="Weekly Team Meeting",
            ...     start_time="2024-01-15T10:00:00",
            ...     end_time="2024-01-15T11:00:00",
            ...     recurrence_type=1,  # Weekly
            ...     interval=1,
            ...     occurrences=10
            ... )
        """
        validate_string_not_empty(subject, "subject")

        try:
            appt = self.application.CreateItem(1)  # 1 = olAppointmentItem

            appt.Subject = subject
            appt.Start = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            appt.End = datetime.fromisoformat(end_time.replace("Z", "+00:00"))

            if location:
                appt.Location = location
            if body:
                appt.Body = body

            # Set recurrence pattern
            rec_pattern = appt.GetRecurrencePattern()
            rec_pattern.RecurrenceType = recurrence_type
            rec_pattern.Interval = interval

            if occurrences:
                rec_pattern.Occurrences = occurrences
            elif end_date:
                rec_pattern.PatternEndDate = datetime.fromisoformat(end_date.replace("Z", "+00:00"))

            appt.Save()

            return dict_to_result(
                success=True,
                message="Recurring appointment created successfully",
                entry_id=appt.EntryID,
                subject=subject,
                recurrence_type=recurrence_type,
                interval=interval,
            )
        except Exception as e:
            raise CalendarOperationError("create_recurring_event", str(e)) from e

    @com_safe("search_appointments")
    def search_appointments(
        self,
        subject: str | None = None,
        location: str | None = None,
        start_date: str | None = None,
        end_date: str | None = None,
        max_results: int = 50,
    ) -> dict[str, Any]:
        """Search for appointments.

        Args:
            subject: Subject text to search for
            location: Location to filter by
            start_date: Start date for search (ISO format)
            end_date: End date for search (ISO format)
            max_results: Maximum number of results

        Returns:
            Dictionary with search results

        Example:
            >>> result = outlook.search_appointments(
            ...     subject="meeting",
            ...     start_date="2024-01-01T00:00:00"
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        calendar = namespace.GetDefaultFolder(9)  # 9 = olFolderCalendar

        try:
            filters = []
            if subject:
                filters.append(f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{subject}%'")
            if location:
                filters.append(
                    f"@SQL=\"http://schemas.microsoft.com/mapi/location\" LIKE '%{location}%'"
                )
            if start_date:
                filters.append(f"@SQL=\"urn:schemas:calendar:dtstart\" >= '{start_date}'")
            if end_date:
                filters.append(f"@SQL=\"urn:schemas:calendar:dtend\" <= '{end_date}'")

            filter_string = " AND ".join(filters) if filters else None

            items = calendar.Items
            items.Sort("[Start]")
            items.IncludeRecurrences = True

            if filter_string:
                items = items.Restrict(filter_string)

            results = []
            for item in items:
                if len(results) >= max_results:
                    break

                results.append(
                    {
                        "entry_id": item.EntryID,
                        "subject": item.Subject,
                        "start_time": str(item.Start),
                        "end_time": str(item.End),
                        "location": item.Location,
                        "is_recurring": item.IsRecurring,
                        "busy_status": item.BusyStatus,
                    }
                )

            return dict_to_result(
                success=True,
                message=f"Found {len(results)} appointment(s)",
                results=results,
                count=len(results),
            )
        except Exception as e:
            raise CalendarOperationError("search_appointments", str(e)) from e

    @com_safe("get_appointments_by_date")
    def get_appointments_by_date(
        self,
        start_date: str,
        end_date: str,
    ) -> dict[str, Any]:
        """Get all appointments in a date range.

        Args:
            start_date: Start date (ISO format)
            end_date: End date (ISO format)

        Returns:
            Dictionary with appointments

        Example:
            >>> result = outlook.get_appointments_by_date(
            ...     start_date="2024-01-15T00:00:00",
            ...     end_date="2024-01-15T23:59:59"
            ... )
        """
        return self.search_appointments(
            start_date=start_date,
            end_date=end_date,
            max_results=100,
        )

    @com_safe("set_reminder")
    def set_reminder(
        self,
        appointment_entry_id: str,
        reminder_minutes: int,
    ) -> dict[str, Any]:
        """Set a reminder for an appointment.

        Args:
            appointment_entry_id: Entry ID of the appointment
            reminder_minutes: Minutes before start to remind

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.set_reminder(
            ...     appointment_entry_id="...",
            ...     reminder_minutes=30
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        try:
            appt.ReminderSet = True
            appt.ReminderMinutesBeforeStart = reminder_minutes
            appt.Save()

            return dict_to_result(
                success=True,
                message="Reminder set successfully",
                subject=appt.Subject,
                reminder_minutes=reminder_minutes,
            )
        except Exception as e:
            raise CalendarOperationError("set_reminder", str(e)) from e

    @com_safe("set_busy_status")
    def set_busy_status(
        self,
        appointment_entry_id: str,
        busy_status: int,  # 0=Free, 1=Tentative, 2=Busy, 3=OutOfOffice
    ) -> dict[str, Any]:
        """Set busy status for an appointment.

        Args:
            appointment_entry_id: Entry ID of the appointment
            busy_status: Busy status (0=Free, 1=Tentative, 2=Busy, 3=OutOfOffice)

        Returns:
            Dictionary with result

        Example:
            >>> result = outlook.set_busy_status(
            ...     appointment_entry_id="...",
            ...     busy_status=3  # Out of Office
            ... )
        """
        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        try:
            appt.BusyStatus = busy_status
            appt.Save()

            status_names = {
                0: "Free",
                1: "Tentative",
                2: "Busy",
                3: "Out of Office",
            }

            return dict_to_result(
                success=True,
                message="Busy status set successfully",
                subject=appt.Subject,
                busy_status=status_names.get(busy_status, "Unknown"),
            )
        except Exception as e:
            raise CalendarOperationError("set_busy_status", str(e)) from e

    @com_safe("export_appointment_ics")
    def export_appointment_ics(
        self,
        appointment_entry_id: str,
        output_path: str,
    ) -> dict[str, Any]:
        """Export an appointment to ICS file.

        Args:
            appointment_entry_id: Entry ID of the appointment
            output_path: Path to save the ICS file

        Returns:
            Dictionary with export result

        Example:
            >>> result = outlook.export_appointment_ics(
            ...     appointment_entry_id="...",
            ...     output_path="C:/exports/meeting.ics"
            ... )
        """
        validate_string_not_empty(output_path, "output_path")

        namespace = self.application.GetNamespace("MAPI")
        appt = namespace.GetItemFromID(appointment_entry_id)

        if appt is None:
            raise OutlookItemNotFoundError("appointment", appointment_entry_id)

        try:
            appt.SaveAs(output_path, 8)  # 8 = olICal format

            return dict_to_result(
                success=True,
                message="Appointment exported successfully",
                subject=appt.Subject,
                output_path=output_path,
            )
        except Exception as e:
            raise CalendarOperationError("export_appointment_ics", str(e)) from e
