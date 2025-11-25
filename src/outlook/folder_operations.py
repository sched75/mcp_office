"""Folder operations mixin for Outlook service.

This module provides folder-related functionality (7 methods).
"""

from typing import Any

from ..core.exceptions import FolderOperationError
from ..utils.com_wrapper import com_safe
from ..utils.helpers import dict_to_result
from ..utils.validators import validate_string_not_empty


class FolderOperationsMixin:
    """Mixin providing folder operations for Outlook.

    Provides 7 methods for managing folders:
    - create_folder
    - delete_folder
    - rename_folder
    - move_folder
    - list_folders
    - get_folder_item_count
    - get_unread_count
    """

    @com_safe("create_folder")
    def create_folder(
        self,
        folder_name: str,
        parent_folder: str = "Inbox",
    ) -> dict[str, Any]:
        """Create a new folder.

        Args:
            folder_name: Name of the new folder
            parent_folder: Parent folder path (default: Inbox)

        Returns:
            Dictionary with creation result

        Raises:
            FolderOperationError: If folder creation fails

        Example:
            >>> result = outlook.create_folder("Projects", parent_folder="Inbox")
        """
        validate_string_not_empty(folder_name, "folder_name")

        namespace = self.application.GetNamespace("MAPI")

        try:
            # Get parent folder
            parent = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if parent_folder != "Inbox":
                for fname in parent_folder.split("/"):
                    parent = parent.Folders(fname)

            # Create new folder
            new_folder = parent.Folders.Add(folder_name)

            return dict_to_result(
                success=True,
                message="Folder created successfully",
                folder_name=folder_name,
                parent_folder=parent_folder,
                folder_path=f"{parent_folder}/{folder_name}",
            )
        except Exception as e:
            raise FolderOperationError("create", folder_name, str(e)) from e

    @com_safe("delete_folder")
    def delete_folder(self, folder_path: str) -> dict[str, Any]:
        """Delete a folder.

        Args:
            folder_path: Path to the folder to delete

        Returns:
            Dictionary with deletion result

        Raises:
            FolderOperationError: If folder deletion fails

        Example:
            >>> result = outlook.delete_folder("Inbox/Old Projects")
        """
        validate_string_not_empty(folder_path, "folder_path")

        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to folder
            folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            parts = folder_path.split("/")

            for part in parts:
                folder = folder.Folders(part)

            folder_name = folder.Name
            folder.Delete()

            return dict_to_result(
                success=True,
                message="Folder deleted successfully",
                folder_path=folder_path,
                folder_name=folder_name,
            )
        except Exception as e:
            raise FolderOperationError("delete", folder_path, str(e)) from e

    @com_safe("rename_folder")
    def rename_folder(self, folder_path: str, new_name: str) -> dict[str, Any]:
        """Rename a folder.

        Args:
            folder_path: Path to the folder to rename
            new_name: New name for the folder

        Returns:
            Dictionary with rename result

        Raises:
            FolderOperationError: If folder rename fails

        Example:
            >>> result = outlook.rename_folder("Inbox/Projects", "Active Projects")
        """
        validate_string_not_empty(folder_path, "folder_path")
        validate_string_not_empty(new_name, "new_name")

        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to folder
            folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            parts = folder_path.split("/")

            for part in parts:
                folder = folder.Folders(part)

            old_name = folder.Name
            folder.Name = new_name

            return dict_to_result(
                success=True,
                message="Folder renamed successfully",
                old_name=old_name,
                new_name=new_name,
                old_path=folder_path,
            )
        except Exception as e:
            raise FolderOperationError("rename", folder_path, str(e)) from e

    @com_safe("move_folder")
    def move_folder(
        self,
        folder_path: str,
        destination_path: str,
    ) -> dict[str, Any]:
        """Move a folder to a different location.

        Args:
            folder_path: Path to the folder to move
            destination_path: Path to destination parent folder

        Returns:
            Dictionary with move result

        Raises:
            FolderOperationError: If folder move fails

        Example:
            >>> result = outlook.move_folder(
            ...     "Inbox/Old",
            ...     "Archive"
            ... )
        """
        validate_string_not_empty(folder_path, "folder_path")
        validate_string_not_empty(destination_path, "destination_path")

        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to source folder
            source_folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            for part in folder_path.split("/"):
                source_folder = source_folder.Folders(part)

            # Navigate to destination folder
            dest_folder = namespace.GetDefaultFolder(6)
            for part in destination_path.split("/"):
                dest_folder = dest_folder.Folders(part)

            folder_name = source_folder.Name
            source_folder.MoveTo(dest_folder)

            return dict_to_result(
                success=True,
                message="Folder moved successfully",
                folder_name=folder_name,
                source=folder_path,
                destination=destination_path,
            )
        except Exception as e:
            raise FolderOperationError("move", folder_path, str(e)) from e

    @com_safe("list_folders")
    def list_folders(
        self,
        parent_folder: str = "Inbox",
        recursive: bool = False,
    ) -> dict[str, Any]:
        """List all folders in a parent folder.

        Args:
            parent_folder: Path to parent folder (default: Inbox)
            recursive: Whether to list folders recursively

        Returns:
            Dictionary with folder list

        Example:
            >>> result = outlook.list_folders("Inbox", recursive=True)
            >>> for folder in result['folders']:
            ...     print(folder['name'])
        """
        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to parent folder
            folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if parent_folder != "Inbox":
                for part in parent_folder.split("/"):
                    folder = folder.Folders(part)

            def list_subfolders(parent, path="", level=0):
                folders = []
                for subfolder in parent.Folders:
                    folder_info = {
                        "name": subfolder.Name,
                        "path": f"{path}/{subfolder.Name}" if path else subfolder.Name,
                        "item_count": subfolder.Items.Count,
                        "unread_count": subfolder.UnReadItemCount,
                        "level": level,
                    }
                    folders.append(folder_info)

                    if recursive:
                        folders.extend(
                            list_subfolders(
                                subfolder,
                                folder_info["path"],
                                level + 1,
                            )
                        )
                return folders

            folders = list_subfolders(folder)

            return dict_to_result(
                success=True,
                message=f"Found {len(folders)} folder(s)",
                parent_folder=parent_folder,
                folders=folders,
                count=len(folders),
            )
        except Exception as e:
            raise FolderOperationError("list", parent_folder, str(e)) from e

    @com_safe("get_folder_item_count")
    def get_folder_item_count(self, folder_path: str = "Inbox") -> dict[str, Any]:
        """Get the number of items in a folder.

        Args:
            folder_path: Path to the folder (default: Inbox)

        Returns:
            Dictionary with item count

        Example:
            >>> result = outlook.get_folder_item_count("Inbox")
            >>> print(result['item_count'])
        """
        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to folder
            folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if folder_path != "Inbox":
                for part in folder_path.split("/"):
                    folder = folder.Folders(part)

            return dict_to_result(
                success=True,
                message="Item count retrieved successfully",
                folder_path=folder_path,
                folder_name=folder.Name,
                item_count=folder.Items.Count,
            )
        except Exception as e:
            raise FolderOperationError("get_count", folder_path, str(e)) from e

    @com_safe("get_unread_count")
    def get_unread_count(self, folder_path: str = "Inbox") -> dict[str, Any]:
        """Get the number of unread items in a folder.

        Args:
            folder_path: Path to the folder (default: Inbox)

        Returns:
            Dictionary with unread count

        Example:
            >>> result = outlook.get_unread_count("Inbox")
            >>> print(result['unread_count'])
        """
        namespace = self.application.GetNamespace("MAPI")

        try:
            # Navigate to folder
            folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
            if folder_path != "Inbox":
                for part in folder_path.split("/"):
                    folder = folder.Folders(part)

            return dict_to_result(
                success=True,
                message="Unread count retrieved successfully",
                folder_path=folder_path,
                folder_name=folder.Name,
                unread_count=folder.UnReadItemCount,
                total_count=folder.Items.Count,
            )
        except Exception as e:
            raise FolderOperationError("get_unread_count", folder_path, str(e)) from e
