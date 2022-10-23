Option Explicit

'************************************************
'*                                              *
'*  More information at:                        *
'*  https://github.com/jonasroslund/multiteams  *
'*                                              *
'************************************************



'***************************************************************
' Define constants
'***************************************************************

Const VERIFICATION_FILE_NAME = "This is a custom Teams folder.txt"


'***************************************************************
' Define support function(s)
'***************************************************************

Function ValidProfileName(Name)
    ' Checks if profile name can be used as a folder name

    Dim BadChars, k
    BadChars = ":\/?*<>|"""

    For k = 1 To Len(BadChars)
        If InStr(Name, Mid(BadChars, k, 1)) > 0 Then
            ValidProfileName = False
            Exit Function
        End If
    Next

    ValidProfileName = True

End Function


Function CreateDeepFolder(FolderName)
    ' Create folder, with parent folders if needed

    Dim FSO, TotalPath, Folder
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If Right(FolderName, 1) = "\" Then FolderName = Left(FolderName, Len(FolderName) - 1)

    TotalPath = ""

    For Each Folder In Split(FolderName, "\")
        
        TotalPath = TotalPath & Folder & "\"

        If FSO.FileExists(TotalPath) Then
            CreateDeepFolder = False
            Exit Function
        ElseIf Not FSO.FolderExists(TotalPath) Then
            FSO.CreateFolder(TotalPath)
        End If
        
    Next

    CreateDeepFolder = True

End Function


Function CheckVerificationFileInFolder(FolderName)
    ' Check if folder has a file verifying it as a custom Teams folder

    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")

    CheckVerificationFileInFolder = FSO.FileExists(FolderName & "\" & VERIFICATION_FILE_NAME)

End Function


Function CreateVerificationFileInFolder(FolderName)
    ' Create a file to identify folder as a custom Teams folder

    Dim FSO, File
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Set File = FSO.CreateTextFile(FolderName & "\" & VERIFICATION_FILE_NAME)
    File.WriteLine("Custom Teams folder")
    File.Close()

End Function


'***************************************************************
' Begin Main Script
'***************************************************************

Dim FSO
Dim Arg, ProfileName, Wsh, Env, Res
Dim MyUserProfile
Dim DefaultCustomFolderLocation, CustomFolder
Dim ScriptName, DotPos

Set Arg = WScript.Arguments
Set FSO = CreateObject("Scripting.FileSystemObject")

Set Wsh = CreateObject("WScript.Shell")
Set Env = Wsh.Environment("Process")

MyUserProfile = Env("USERPROFILE")
DefaultCustomFolderLocation = MyUserProfile & "\AppData\Local\Microsoft\Teams\CustomProfiles"


'----------------
' Check argument
'----------------

If Arg.Count > 0 Then
    
    If InStr(Arg(0), "\") = 0 Then
        ' Argument is a profile name

        ProfileName = Arg(0)

        ' Validate profile name, to be used as a folder name
        If Not ValidProfileName(ProfileName) Then
            MsgBox "Invalid profile name:" & vbNewLine & ProfileName
            WScript.Quit
        End If

        CustomFolder = DefaultCustomFolderLocation & "\" & ProfileName

    Else
        ' Argument is a folder path

        CustomFolder = Arg(0)

    End If

Else
    ' Use script file name as profile name

    ScriptName = WScript.ScriptName
    DotPos = InStrRev(ScriptName, ".")
    If DotPos = 0 Then
        ProfileName = ScriptName
    Else
        ProfileName = Left(ScriptName, DotPos - 1)
    End If

    CustomFolder = DefaultCustomFolderLocation & "\" & ProfileName

End If

'------------------------------------------------------
' Check if folder is suitable as a custom Teams folder
'------------------------------------------------------

If FSO.FolderExists(CustomFolder) Then
    ' Folder does not exist

    If Not CheckVerificationFileInFolder(CustomFolder) Then
        ' Verification file does not exist

        ' Get confirmation from user to continue anyway
        Res = MsgBox("Folder exists, but does not seem to contain custom Teams data." & vbNewLine & _
            "Continue anyway?", vbOKCancel)

        If Not Res = vbOK Then
            WScript.Quit
        End If

        ' Create verification file
        CreateVerificationFileInFolder(CustomFolder)

    End If
Else

    Res = MsgBox("Folder does not exist." & vbNewLine & _
        "Continue setting up new custom Teams folder?", vbOKCancel)

    If Not Res = vbOK Then
        WScript.Quit
    End If

    ' Create folder and verification file
    CreateDeepFolder(CustomFolder)
    CreateVerificationFileInFolder(CustomFolder)

End If

'----------------------
' Set up and run Teams
'----------------------

' Set fake user profile folder, to trick Teams
Env("USERPROFILE") = CustomFolder 

' Set working directory
Wsh.CurrentDirectory = MyUserProfile & "\AppData\Local\Microsoft\Teams"

' Start Teams, initiating data if it doesn't exist
Wsh.Run("""" & MyUserProfile & "\AppData\Local\Microsoft\Teams\Update.exe"" --processStart ""Teams.exe""")
