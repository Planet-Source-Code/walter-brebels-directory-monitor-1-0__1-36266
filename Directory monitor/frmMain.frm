VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Directory Monitor 1.0"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   462
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cmmd 
      Left            =   0
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox txtFolder 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Text            =   "C:\"
      Top             =   5205
      Width           =   3495
   End
   Begin MSComctlLib.ListView lstFiles 
      Height          =   4575
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   8070
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Action"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date last changed"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start Monitoring"
      Height          =   375
      Left            =   5340
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   5520
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Directory monitor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   6675
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000018&
      BackStyle       =   1  'Opaque
      Height          =   5535
      Left            =   120
      Top             =   120
      Width           =   6690
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuSave 
         Caption         =   "Save list..."
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear list"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*********************************************
'**                                         **
'**       Created by Walter Brebels         **
'**         Walter.Brebels@gmx.net          **
'**---------------------------------------  **
'** Date: 23/06/2002                        **
'** Function: monitors a directory for files**
'**           being: created, modified,     **
'**           removed and renamed using FSO **
'** Copyrights: you can use/modify this code**
'** freely in your appications without      **
'** mentioning my name (aldo it would be    **
'** appreciated).                           **
'**-----------------------------------------**
'** For more information about the FSO:     **
'** http://msdn.microsoft.com/library/      **
'** default.asp?url=/library/en-us/script56/**
'** html/jsfsotutor.asp                     **
'**                                         **
'*********************************************

'FileSystemObject declarations
Dim objFSO, objFolder, sFile
'the files that are currently in the directory
Dim Files As Collection, oldFiles As Collection

Dim LastFileCount As Long
Dim newFile As ListItem
'Just a variable for a For...Next loop
Dim Object As Variant
'the folder wich we are monitoring
Dim sFolder As String

Private Sub cmdBrowse_Click()
Dim FolderName As String
FolderName = GetFolderName(hWnd, "Choose a directory to monitor")
If FolderName <> "" Then
    txtFolder.Text = FolderName
End If
End Sub

Private Sub cmdStart_Click()
If Timer1.Enabled = True Then
    Timer1.Enabled = False
    cmdStart.Caption = "Start Monitoring"
Else
    Timer1.Enabled = True
    cmdStart.Caption = "Stop Monitoring"
End If
End Sub


Private Sub Form_Load()
'set the new collection
Set Files = New Collection
Set oldFiles = New Collection
'setting default folder
sFolder = "C:\"
'creating our FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.getfolder(sFolder)

'Set the Attributes of each file to zero (meaning to: Normal)
'and getting each file
For Each sFile In objFolder.Files
    sFile.Attributes = 0
    Files.Add CStr(sFile), CStr(sFile)
Next sFile

LastFileCount = Files.Count
Set oldFiles = Files
Label1.Caption = "Directory monitor - Monitoring: " & objFolder
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopupMenu mnuPopUp
End If
End Sub

Private Sub mnuClear_Click()
lstFiles.ListItems.Clear
End Sub

Private Sub mnuSave_Click()
'basic options for the commondialog
With Cmmd
    .DialogTitle = "Choose a file"
    .Filter = "Text file (*.txt)|*.txt"
    .ShowSave
    If .FileName = "" Then Exit Sub
    If FileExist(.FileName) Then
        If MsgBox("File exist, do you want to overwrite it?", vbQuestion + vbYesNo, "Direcotry monitor 1.0") = vbYes Then
        SaveList .FileName
        End If
    Else
    SaveList .FileName
    End If
End With
End Sub

Private Sub Timer1_Timer()
If Not objFSO.folderexists(sFolder) Then
    MsgBox "Folder has been removed, please choose another folder", vbInformation + vbOKOnly + vbMsgBoxSetForeground, "Directory monitor 1.0"
    cmdStart_Click
End If

Set objFolder = objFSO.getfolder(sFolder)
Set Files = New Collection
'Going thru each file in the directory
For Each sFile In objFolder.Files
    
    'Add file to the collection
    Files.Add CStr(sFile), CStr(sFile)
            
    'Checking if there's a new file
    If CheckEntry(oldFiles, CStr(sFile)) = False And CheckDate(CStr(sFile.datecreated)) = True Then
        AddFile sFile, "New File"
        sFile.Attributes = 0
    End If
    'Checking if the file name just changed
    If CheckEntry(oldFiles, CStr(sFile)) = False And CheckDate(CStr(sFile.datecreated)) = False Then
        AddFile sFile, "Filename changed"
        sFile.Attributes = 0
    End If
    
    'Checking the attributes (32 = file modified)
    If sFile.Attributes <> 0 Then
        AddFile sFile, "File changed"
        sFile.Attributes = 0
    End If
Next sFile

'Checking for deleted files
If Files.Count < LastFileCount Then
    For Each Object In oldFiles
    If FileExist(CStr(Object)) = False Then AddFile CStr(Object), "File Deleted"
    Next Object
End If

LastFileCount = Files.Count
Set oldFiles = Files
End Sub

Private Sub AddFile(File As Variant, Action As String)
On Error GoTo ErrOc
Set newFile = lstFiles.ListItems.Add(1, , File)
newFile.SubItems(1) = Action
newFile.SubItems(2) = Date & " - " & Time
Exit Sub
ErrOc:
MsgBox File
End Sub

Private Function FileExist(Path As String) As Boolean
Dim CheckFile As String
CheckFile = Dir(Path)
FileExist = IIf(CheckFile <> "", True, False)
End Function

Private Function CheckEntry(cColl As Collection, sEntry As String) As Boolean
On Error GoTo ErrHandler
Dim var
var = cColl.Item(sEntry)
CheckEntry = True
Exit Function

ErrHandler:
CheckEntry = False
End Function

'Function to check the difference between a certain date and now
'USE: to see if file has just been created or if the file name has
'     just been changed.
Private Function CheckDate(sDate As String) As Boolean
Dim Parts() As String, TimeParts() As String
Dim Second1 As Integer, Second2 As Integer
Parts = Split(sDate, " ", , vbTextCompare)
If UBound(Parts) <> 1 Then Exit Function
TimeParts = Split(Parts(1), ":", , vbTextCompare)
Second1 = CInt(TimeParts(2))
Second2 = Second(Time)
If CDate(Parts(0)) = Date And Abs(Second1 - Second2) <= 2 Then
    CheckDate = True
Else
    CheckDate = False
End If
End Function

Private Sub txtFolder_KeyPress(KeyAscii As Integer)
'if the user presses enter
If KeyAscii = 13 Then
    'checking if folder exists
    If objFSO.folderexists(txtFolder.Text) Then
    sFolder = txtFolder.Text
    'SAME ROUTINE AS ON STARTUP:
    'Set the Attributes of each file to zero (meaning to: Normal)
    'and getting each file
    Set objFolder = objFSO.getfolder(sFolder)
    For Each sFile In objFolder.Files
        sFile.Attributes = 0
        Files.Add CStr(sFile), CStr(sFile)
    Next sFile
    LastFileCount = Files.Count
    Set oldFiles = Files
    Label1.Caption = "Directory monitor - Monitoring: " & objFolder
    Else
    MsgBox "Folder does not exists!", vbInformation, App.Title
    End If
Else
End If
End Sub

Private Sub SaveList(Path As String)
Dim FF As Long, lstItem As ListItem
'Asign an unused filenumber
FF = FreeFile
'If file exists then remove it
If FileExist(Path) Then Kill Path
'Open the file For Output (means we're going to write to it)
Open Path For Output As #FF
    Print #FF, "        *** Directory monitor 1.0 ***"
    'just print some general and usefull information ;)
    Print #FF, "Created by Walter Brebels - Walter.Brebels@gmx.net"
    Print #FF,
        'go thru each item in the listview
        For Each lstItem In lstFiles.ListItems
        Print #FF, lstItem.Text & vbTab & lstItem.SubItems(1) & vbTab & lstItem.SubItems(2)
        Next lstItem
    Print #FF,
    Print #FF, "              *** End list ***"
'Close the file
Close FF
End Sub
