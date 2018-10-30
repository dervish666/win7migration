# win7migration

Backup and restore scripts to assist with win 7 to 10 migration

List of files:

    BackupUserData.ps1
        This script will backup the users data to the destination of your choice. Recommend backing up to USB key. 

    RestoreUserData.ps1 
        This will restore all the user data and copy their docs etc to onedrive

    CheckInstalled.ps1
        This will check if the essential software is a) installed and b) running. 

    defaults.txt
        This is for your list of default values for things like the backup destination, it also includes the default 
        answers so the scipt can be run unattended. 

Each of the .ps1 files includes a batch file to facilitate running easily. The script can be run from the batch or directly in powershell if script execution is enabled. 

Process:
    This is the overall process for migrating users 

    Backup Process 

        Each market adjusts the defaults.txt file to suit their usage
        Format USB stick to NTFS with at least 64GB capacity
        Copy entire win7migration folder to usb stick
        Floorwalker then takes USB to the users machine runs BackupUserData.bat from the USB
        Floorwalker chooses USB drive letter, lets script finish
        Data is now on USB and optionally backed up to the NAS

    Restore Process

        Adjust defaults.txt
        Using stick from backup, run the Restore script
        If first time running script, do a gpupdate
        Reboot
        Check Onedrive is configured, if not wait
        Run restore script
        Enter users old name, allow script to finish
        Open Outlook, import any PSTs from the Downloads folder


Backup User Data script. 
    A bit more detail on the individual scripts.

Process:

    Backup user data to a secure location 
    Backup the my docs/desktop/pictures/videos folders
    Backup Google Chrome bookmarks
    Find all the active PSTs on the computer
    Back up all active PSTs
    Disconnect PSTs from outlook
    Further back up all the users data to a secondary location
    Check for any other PSTs on the users hard drive


Restore User Data script.

Process:

    Check the new computer for essential software
    Check all necessary processes are running
    Ensure GPUdate is run to trigger OneDrive install
    Special check for Office 13 for certain markets
    Restore the users folders
    Convert Sticky Notes for use with Win 10
    Restore Chrome bookmarks
    Restore my docs/desktop/pictures/videos folders to OneDrive
    Copy the PST files to Downloads folder for ease of importing to O365


defaults text file
    Looks a bit like this: 

        nas:\\10.61.11.125\ukipst\Import
        def:\\UKBHSR255\PSTImport
        creds:ukiadm
        UseBackupDefaults:true
        AlwaysBackupToUSB:true
        AlwaysBackupDocs:true
        AlwaysBackupPST:true
        AlwaysBackupToNAS:false
        AlwaysCheckForPSTs:true
        UseRestoreDefaults:true
        AlwaysRestoreFromUSB:true
        AlwaysRestoreUsersFolders:true
        AlwaysRestoreToOneDrive:true
        AlwaysCopyPSTSToDLs:true

This file will set the defaults for the script, the left hand side is the variable name and the right hand side is the value. Both scripts will read from the same file but only use the variable they need. 

                                DO NOT CHANGE ANY VALUE ON THE LEFT HAND SIDE!!!

Each entry in detail: 

    Universal variables:

    nas: - This is a text string for the secondary backup location 
    def: - This is the alternative backup (It is recommended to backup to USB for speed)
    creds: - The username for the secondary location
    
    Backup variables:
    
    UseBackupDefaults: - This is a flag to use the defaults, if set to false the backup script will ask for each section
    AlwaysBackupToUSB: - Sets the script to backup to USB, it will still ask where the drive is located
    AlwaysBackupDocs: - Sets the script to backup the My Documents/Desktop/Pictures/Videos folders
    AlwaysBackupPST: - Set the script to check for active PSTs and back them up
    AlwaysBackupToNAS: - Sets the script to backup the backup to the secondary location
    AlwaysCheckForPSTs: - Sets the script to check the hard drive for any inactive PSTs

    Restore variables:

    UseRestoreDefaults: - Same as backup but for the restore script
    AlwaysRestoreFromUSB: - Sets the script to restore from USB
    AlwaysRestoreUsersFolders: - Sets the script to restore the Sticky Notes, Signatures, etc 
    AlwaysRestoreToOneDrive: - Sets the script to always restore the My Docs and Desktop folders to OneDrive
    AlwaysCopyPSTSToDLs: - Sets the script to copy all PSTs into the Download folder 