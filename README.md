# MultiTeams
This script allows you to use multiple sessions of Microsoft Teams simultaneously. Each session can be logged in with different accounts or guest accounts. Some use cases are:
* You have been invited as a guest user in business partner's team and want to access both the business partner's team as guest and your company's teams simultaneously.
* You have multiple Microsoft 365 accounts and need to be logged in to several Teams sessions simultaneously.

Separate folders will be used to store data for the additional Microsoft Teams sessions.

# Usage
## Preparation
1. Download [the repository from GitHub](https://github.com/jonasroslund/multiteams/archive/refs/heads/main.zip) and extract `MultiTeams.vbs` to your desktop.
2. Rename `MultiTeams.vbs` to describe the additional setup, e.g. `MyOtherCompanyTeams.vbs`.

## Runtime
1. Start Microsoft Teams as you normally do. This will open your default setup .
2. Double click the script you saved to your desktop. This will open a new Microsoft Teams window.
3. In the new Microsoft Teams window, change to the account or guest account you want to access (e.g. a client). This will be used next time you click the script.
4. Congratulations, you can now use multiple Microsoft Teams accounts simultaneously.

## Need more simultaneous sessions?
If you need more simultaneous sessions, just save more copies of the script with different names.

# How it works
MultiTeams.vbs creates a fake user profile folder. When Microsoft Teams starts, it will try to find its data in the fake profile folder. If the profile is empty, Microsoft Teams will set up a complete data structure to store data and start Microsoft Teams. If data exists from a previous session, Microsoft Teams will start a new session based on the fake profile folder.

Calling the script without any argument/parameter will store data in a folder based on the script name:
`C:\Users\<YourUserName>\AppData\Local\Microsoft\Teams\CustomProfiles\<ScriptName>`.

To get more control of where to store data, see [Advanced usage](#advanced-usage).

# Advanced usage
If you need more control over the name of the folder where the session is stored, you can call the script using one of alternatives below. You don't need to change the script name in the preparation phase for the advanced usage, as the folder name will be decided by the script argument.

## Advanced usage 1. Folder name as argument

```
MultiTeams.vbs MySecondCompany
```
This will store session data in `C:\Users\<YourUserName>\AppData\Local\Microsoft\Teams\CustomProfiles\MySecondCompany`.

Tip: Store the script in a folder of your choice and create a desktop shortcut with a reference to the script with the folder name as an argument.

## Advanced usage 2. Full path as argument
```
MultiTeams.vbs C:\MyTeamsData\MySecondCompany
```
This will store session data in `C:\MyTeamsData\MySecondCompany`.

Tip: Store the script in a folder of your choice and create a desktop shortcut with a reference to the script with the full path as an argument.

