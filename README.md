# MultiTeams
This script allows you to use multiple sessions of Microsoft Teams simultaneously. Each session can be logged in with different accounts or guest accounts. Some use cases are:
* You have been invited as a guest user in business partner's team and want access both the business partner's team and your company's teams simultaneously.
* You have multiple Microsoft 365 accounts and need to be logged in to several Teams sessions simultaneously.

Separate folders will be used to store data for the additional Microsoft Teams sessions.

# Usage
1. Start your normal Teams session.
2. Run MultiTeams.vbs using one or more of the alternatives below.

## Alternative 1. Use script name as folder name

This is probably the easiest way to use the script.

Calling the script without any argument will store data in a folder based on the script name. This allows you to put a copy of the script on your desktop, rename it and double click it to start a custom session of Microsoft Teams.

Example:
1. Download the script `MultiTeams.vbs` to your desktop.
2. Rename it to `MyOtherCompanyTeams.vbs`.
3. Double click the file.
4. A separate session of Teams will be started. Session data will be stored in `C:\Users\<YourUserName>\AppData\Local\Microsoft\Teams\CustomProfiles\MyOtherCompanyTeams`

If you need more simultaneous sessions, just save more copies of the script with different names.

## Alternative 2. Folder name as argument

```
MultiTeams.vbs MySecondCompany
```
This will store session data in `C:\Users\<YourUserName>\AppData\Local\Microsoft\Teams\CustomProfiles\MySecondCompany`

Tip: Store the script in a folder of your choice and create a desktop shortcut with a reference to the script with the folder name as an argument.

## Alternative 3. Full path as argument
```
MultiTeams.vbs C:\MyTeamsData\MySecondCompany
```
This will store session data in `C:\MyTeamsData\MySecondCompany`

Tip: Store the script in a folder of your choice and create a desktop shortcut with a reference to the script with the full path as an argument.

# How it works
MultiTeams.vbs creates a fake user profile folder. When Microsoft Teams starts, it will try to find its data in the fake profile folder. If the profile is empty, Microsoft Teams will set up a complete data structure to store data and start Microsoft Teams. If data exists from a previous session, Microsoft Teams will start a new session based on the fake profile folder.
