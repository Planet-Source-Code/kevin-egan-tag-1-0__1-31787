VERSION 5.00
Begin VB.Form Tagfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TAG 1.0"
   ClientHeight    =   3390
   ClientLeft      =   6000
   ClientTop       =   5190
   ClientWidth     =   8310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   8310
   Begin VB.CommandButton CmdRem 
      Caption         =   "Remove Tag"
      Height          =   375
      Left            =   5880
      TabIndex        =   18
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add Tag"
      Height          =   375
      Left            =   4680
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   16
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Select Tag"
      Height          =   2175
      Left            =   6600
      TabIndex        =   9
      ToolTipText     =   "Select the Tag you wish to Add or Remove"
      Top             =   360
      Width           =   1500
      Begin VB.TextBox TxtTag 
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.OptionButton optSPECIFYTAG 
         Caption         =   "Specify"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton optES 
         Caption         =   "_es"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   615
      End
      Begin VB.OptionButton optFR 
         Caption         =   "_fr"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optDE 
         Caption         =   "_de"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select file type"
      Height          =   2175
      Left            =   4800
      TabIndex        =   3
      ToolTipText     =   "Select the type of file you wish to change"
      Top             =   360
      Width           =   1500
      Begin VB.OptionButton optSPECIFYFILE 
         Caption         =   "Specify"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton optGIF 
         Caption         =   "gif"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optJPG 
         Caption         =   "jpg"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton OptHTML 
         Caption         =   "html"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox TxtFile 
         Enabled         =   0   'False
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   2280
      TabIndex        =   2
      Top             =   480
      Width           =   2175
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "File List:"
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu cmdfile 
      Caption         =   "&File"
      Begin VB.Menu MENUcmdaddtag 
         Caption         =   "&Add Tag"
         Shortcut        =   ^A
      End
      Begin VB.Menu MENUcmdremovetag 
         Caption         =   "&Remove Tag"
         Shortcut        =   ^R
      End
      Begin VB.Menu bar 
         Caption         =   "-"
      End
      Begin VB.Menu cmdexit 
         Caption         =   "&Exit"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu cmdhelp 
      Caption         =   "&Help"
      Begin VB.Menu cmdtag 
         Caption         =   "&Tag Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu bar2 
         Caption         =   "-"
      End
      Begin VB.Menu cmdabout 
         Caption         =   "&About Tag"
      End
   End
End
Attribute VB_Name = "Tagfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim directory As String
Dim fName As String
'dim fieldtype as Global s
'global filetype As String
'Static tagname As String
Dim tagname As String
Dim filetype As String
Dim Flist(1 To 300) As String

Sub getfiletype()


'/////get file type


filetype = ""
If OptHTML = True Then filetype = ".html"
If optJPG = True Then filetype = ".jpg"
If optGIF = True Then filetype = ".gif"
If optSPECIFYFILE = True Then

Dim str As String
    str = TxtFile.Text
    If Left(str, 1) = "." Then
        filetype = str
    Else
        filetype = "." & str
    End If
End If
'///////////

End Sub

Sub getTagname()

'///////get tag


tagname = ""
If optFR = True Then tagname = "_fr"
If optDE = True Then tagname = "_de"
If optES = True Then tagname = "_es"
If optSPECIFYTAG = True Then
Dim str As String
    str = TxtTag.Text
    If Left(str, 1) = "_" Then
        filetype = str
    Else
        filetype = "_" & str
    End If
End If

'/////////

End Sub



Private Sub cmdabout_Click()
aboutfrm.Show (1)
End Sub

Private Sub CmdAdd_Click()
If optSPECIFYFILE = True Then
    If TxtFile.Text = "" Then
    MsgBox "You have not filled in the specify box on the 'Select file type' menu"
    Exit Sub
    End If
End If

If optSPECIFYTAG = True Then
    If TxtTag.Text = "" Then
    MsgBox "You have not filled in the specify box on the 'Select tag' menu"
    Exit Sub
    End If
End If

If OptHTML.Value = 0 Then
    If optJPG.Value = 0 Then
    If optGIF.Value = 0 Then
    If optSPECIFYFILE.Value = 0 Then
   
    MsgBox "You have not selected a file type"
    Exit Sub
End If
End If
End If
End If



If optFR.Value = 0 Then
    If optDE.Value = 0 Then
    If optES.Value = 0 Then
    If optSPECIFYTAG.Value = 0 Then
    MsgBox "You have not selected a tag"
    Exit Sub
End If
End If
End If
End If

Dim i As Integer
Dim extLen As Integer
Dim counter As Integer
Dim tempfName As String
Dim ext As String
Dim fileID As String
Dim withoutTag As String
Dim LenwithoutTag As Integer
Dim looper As Integer

getTagname
getfiletype
getList
i = 1
looper = 1

Do While Flist(looper) <> "" ' Start loop through directory.
    counter = 0
    fName = Flist(looper)
       For i = Len(fName) To 1 Step -1
            counter = counter + 1
            If Mid(fName, i, 1) = "." Then
                ext = Right(fName, counter)
                tempfName = Len(fName) - counter
                fileID = Left(fName, tempfName)
                If Right(fileID, 3) <> tagname Then
                    If Dir$(directory & "\" & fileID & tagname & ext) = "" Then ' see if file exists
                        Name directory & "\" & fName As directory & "\" & fileID & tagname & ext ' add tag
                    End If
                End If
                Exit For
            End If
        Next
    looper = looper + 1
Loop ' end loop through directory.
refreshFiles
End Sub

Private Sub cmdend_Click()
M = MsgBox("Are you sure you want to exit Tag?", vbYesNo + vbQuestion, "Exit TAG?")
If M = vbYes Then End
If M = vbNo Then Exit Sub
End Sub

Private Sub cmdexit_Click()
cmdend.Value = True
End Sub

Private Sub CmdRem_Click()
If optSPECIFYFILE = True Then
    If TxtFile.Text = "" Then
    MsgBox "You have not filled in the specify box on the 'Select file type' menu"
    Exit Sub
    End If
End If
If optSPECIFYTAG = True Then
    If TxtTag.Text = "" Then
    MsgBox "You have not filled in the specify box on the 'Select tag' menu"
    Exit Sub
    End If
End If

If OptHTML.Value = 0 Then
    If optJPG.Value = 0 Then
    If optGIF.Value = 0 Then
    If optSPECIFYFILE.Value = 0 Then
    MsgBox "You have not selected a file type"
    Exit Sub
End If
End If
End If
End If

If optFR.Value = 0 Then
     If optDE.Value = 0 Then
    If optES.Value = 0 Then
    If optSPECIFYTAG.Value = 0 Then
    MsgBox "You have not selected a tag"
    Exit Sub
End If
End If
End If
End If

Dim i As Integer
Dim extLen As Integer
Dim counter As Integer
Dim tempfName As String
Dim ext As String
Dim fileID As String
Dim withoutTag As String
Dim LenwithoutTag As Integer
Dim looper As Integer

getTagname
getfiletype
getList
i = 1
looper = 1

Do While Flist(looper) <> "" ' Start loop through directory.
    counter = 0
    fName = Flist(looper)
    
    
       For i = Len(fName) To 1 Step -1
            counter = counter + 1
            If Mid(fName, i, 1) = "." Then
                ext = Right(fName, counter)
                tempfName = Len(fName) - counter
                fileID = Left(fName, tempfName)
                If Right(fileID, 3) = tagname Then
                    LenwithoutTag = Len(fileID) - 3
                    withoutTag = Left(fileID, LenwithoutTag)
                    If Dir$(directory & "\" & withoutTag & ext) = "" Then ' see if file exists
                        Name directory & "\" & fName As directory & "\" & withoutTag & ext ' rem tag
                    End If
                End If
                Exit For
            End If
        Next
    looper = looper + 1
Loop ' end loop through directory.

refreshFiles
End Sub

Sub refreshFiles()

Dim directory1 As String
List1.Clear
directory = Dir1.Path ' Get directory from listbox
fName = Dir(directory & "\*.*")    ' Retrieve the first entry.

Do While fName <> "" ' Start loop through directory.
    If Right(fName, 3) <> "sys" Then
    If GetAttr(directory & "\" & fName) <> vbDirectory Then '
    ' MsgBox ""
     List1.AddItem fName
    End If
    End If
    fName = Dir
Loop

End Sub



Private Sub cmdtag_Click()
    Dim RetVal
    Dim browserstring As String
    Dim launchbrowser As String
    
    
    browserstring = "rundll32.exe url.dll,FileProtocolHandler "
    
    launchbrowser = browserstring + App.Path + "\help.html"
    
    RetVal = Shell(launchbrowser, 1)
End Sub

Sub getList()

directory = Dir1.Path ' Get directory from listbox
tmpStrg = Dir$(directory & "\*" & filetype)   ' Retrieve the first entry.

counter = 1
    
   If tmpStrg <> "" Then 'have files in the directory
         fName = tmpStrg
         Flist(1) = fName
         tmpStrg = Dir$ 'Go back to the directory to add more
            While Len(tmpStrg) > 0 'While there is still more unadded
                counter = counter + 1
                Flist(counter) = tmpStrg
                tmpStrg = Dir$ 'Go back to the directory to add more
            Wend
    End If
    
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Command2_Click()




End Sub

Private Sub Dir1_Change()
refreshFiles
End Sub

Private Sub Drive1_Change()
On Error GoTo drivenotready
Drive = Drive1.Drive
Dir1.Path = Drive1.Drive
Exit Sub
drivenotready:
MsgBox "Drive Not Ready", vbOKOnly, " "
Drive1.Drive = "c:\"
End Sub

Private Sub Form_Load()
refreshFiles
End Sub




Private Sub MENUcmdaddtag_Click()
CmdAdd.Value = True
End Sub

Private Sub MENUcmdremovetag_Click()
CmdRem.Value = True
End Sub

Private Sub optDE_Click()
TxtTag.Enabled = False
End Sub

Private Sub optES_Click()
TxtTag.Enabled = False
End Sub

Private Sub optFR_Click()
TxtTag.Enabled = False
End Sub

Private Sub optGIF_Click()
TxtFile.Enabled = False
End Sub

Private Sub OptHTML_Click()
TxtFile.Enabled = False
End Sub

Private Sub optJPG_Click()
TxtFile.Enabled = False
End Sub

Private Sub optSPECIFYFILE_Click()
TxtFile.Enabled = True
TxtFile.SetFocus
End Sub

Private Sub optSPECIFYTAG_Click()
TxtTag.Enabled = True
TxtTag.SetFocus
End Sub
