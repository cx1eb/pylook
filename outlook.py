import time
import sys
import os

import win32com.client
import pythoncom

from .enums import ComError, ExchangeStoreType

class OutlookClient:
    def __init__(self, debug=False):
        self.debug = debug
        self.app = None
        self.namespace = None
        self.was_already_running = None

    def _log(self, msg):
        if self.debug:
            print(f'[PYO]: {msg}')

    def _is_outlook_running(self):
        try:
            win32com.client.GetActiveObject("Outlook.Application")
            return True
        except pythoncom.com_error as e:
            hresult = e.args[0]
            err = ComError.from_code(hresult)
            if err is not None and err == ComError.MK_E_UNAVAILABLE:
                self._log(f'{err.name}: Outlook object not found...')
            elif err is not None:
                self._log(f'{err.name}')
            else:
                self._log(f'Unknown COM error: {hresult}')

            return False    
    
    def _force_load(self, folder):
        """Force Outlook to populate child folders for shared mailboxes."""
        try:
            _ = folder.Folders.Count          # touch count
            for i in range(1, folder.Folders.Count + 1):
                sub = folder.Folders.Item(i)  # touch each one
        except Exception:
            pass



    def connect(self):
        """Connect to running outlook COM istance, or start one"""
        check = self._is_outlook_running()
        if check == True:
            self._log('Outlook already running, connecting to instance.')
        else:
            self._log('Starting new outlook connection.')

        try:
            self.app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.app.GetNamespace("MAPI")
        
            version = self.app.Version 
            account_count = self.namespace.Accounts.Count 

            self._log(f'Outlok version: {version}')
            self._log(f'Accounts available: {account_count}')

            return True

        except pythoncom.com_error as e:
            self._log(f'Failed to connect to outlook: {e}')
            self.app = None
            self.namespace = None
            return False

        
    def get_shared_mailboxes(self):
        if not self.is_connected:
            self._log('Outlook connection required in order to list mailboxes...')
            return

        shared = []

        '''
        ExchangeStoreType values:
        0 = Primary mailbox
        1 = Delegate mailbox
        2 = Public folder
        3 = Additional (auto-mapped) mailbox
        4 = Not exchange (PST, etc.)
        5 = Exchange public folder
        '''

        SHARED_TYPES = (ExchangeStoreType.DELEGATE, ExchangeStoreType.ADDITIONAL) #delagate and additional

        for store in self.namespace.Stores:
            try:
                store_type = store.ExchangeStoreType
            except Exception:
                store_type = None

            if store_type in SHARED_TYPES:
                shared.append({
                    "name": store.DisplayName,
                    "smtp": getattr(store, "SmtpAddress", None),
                    "store_id": store.StoreID,
                    "type": store_type,
                })

                #self._log(f"Found shared mailbox: {store.DisplayName}")

        
        return shared

    def get_mailbox_root(self, mailbox_name: str):
        """Gets the root folder of a shared or primary mailbox via its display name"""
        if not self.is_connected:
            self._log('Outlook connection required in order to list mailboxes...')
            return []

        for store in self.namespace.Stores:
            if store.DisplayName.lower() == mailbox_name.lower():  #lower to make sure they match, dont gotta be case senstive
                return store.GetRootFolder()


    def list_folders(self, folder):
        """Returns the subfolders inside a folder."""
        return [folder.Folders.Item(i) for i in range(1, folder.Folders.Count + 1)]


    def walk_folders(self, folder, depth=0):
        """
        Gets every folder in folder under the given folder

        returns tuples of (folder, depth) so you can indent / tack nesting.

        Keeping it a basic return so user can display how they would like
        """
        yield folder, depth

        for sub in folder.Folders:
            yield from self.walk_folders(sub, depth + 1)


    def get_folder_by_path(self, root, path, sep: str = "/"):
        """
        Navigate a folder path inside a mailbox

        Accepts:
            - list/tuple of names (recommended)
            - string with chosen separator (default is /)
        """
        if isinstance(path, (list, tuple)):
            parts = list(path)
        elif isinstance(path, str):
            parts = [p for p in path.split(sep) if p]
        else:
            raise TypeError("path must be a string or list of folder names")

        current = root
        for part in parts:
            found = None
            for i in range(1, current.Folders.Count + 1):
                sub = current.Folders.Item(i)
                if sub.Name.lower() == part.lower():
                    found = sub
                    break
            if found is None:
                raise ValueError(f"Folder '{part}' not found under '{current.Name}'")
            current = found
            self._log(f"Entered folder: {current.Name}")
        return current


    
    def get_shared_inbox(self, smtp_address: str):
        """Get a shared mailbox's Inbox via resolved Recipient (bypasses cache issues)."""
        if not self.is_connected:
            self._log('Outlook connection required in order to list mailboxes...')
            return


        recipient = self.namespace.CreateRecipient(smtp_address)
        recipient.Resolve()

        if not recipient.Resolved:
            raise ValueError(f"Could not resolve recipient: {smtp_address}")

        # 6 = olFolderInbox
        inbox = self.namespace.GetSharedDefaultFolder(recipient, 6)
        return inbox


    def count_emails(self, folder, include_subfolders: bool = False) -> int:
        """
        Count emails in a folder.

        Args:
            folder: Outlook folder object
            include_subfolders: if True, recursively counts items in subfolders too

        Returns:
            int: number of items
        """
        if folder is None:
            return 0

        total = folder.Items.Count

        if include_subfolders:
            for i in range(1, folder.Folders.Count + 1):
                total += self.count_emails(folder.Folders.Item(i), include_subfolders=True)

        return total

    
    
    def copy_email(self, item, target_folder):
        """
        Copy a single email/item into the target folder.

        Args:
            item: Outlook MailItem (or any moveable item)
            target_folder: destination Folder object

        Returns:
            The copied item in the target folder.
        """
        if item is None or target_folder is None:
            raise ValueError("Both item and target_folder are required.")

        copied = item.Copy()         #makes a copy (still in original folder)
        moved = copied.Move(target_folder)  #move the copy to destination
        return moved






    @property
    def version(self):
        return self.app.Version if self.app else None

    @property
    def name(self):
        return self.app.Name if self.app else None
    
    @property
    def is_connected(self):
        return self.app is not None

    @property
    def account_count(self):
        return self.namespace.Accounts.Count if self.namespace else 0

    @property
    def default_profile(self):
        return self.namespace.CurrentProfileName if self.namespace else None
