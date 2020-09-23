VERSION 5.00
Begin VB.Form frmCheckforFile 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Check for File"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   5400
   FillColor       =   &H00808080&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Check4LDB.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUpdateTo 
      Height          =   285
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   10
      Text            =   "s:\quality\database files\quality_be.ldb"
      Top             =   3000
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.TextBox txtUpdateWith 
      Height          =   285
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   9
      Text            =   "s:\quality\database files\quality_be.ldb"
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CheckBox chkUpdateFile 
      BackColor       =   &H80000005&
      Caption         =   "Upload a new file?"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtNumberofMinutes 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   960
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3960
      Top             =   -120
   End
   Begin VB.TextBox txtNumberChecked 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1440
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   120
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Text            =   "s:\quality\database files\quality_be.ldb"
      Top             =   480
      Width           =   4215
   End
   Begin VB.CommandButton cmdBeginCheck 
      Caption         =   "Begin Checking"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label lblUpdateWith 
      BackColor       =   &H80000005&
      Caption         =   "Source File"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lblUpdateto 
      BackColor       =   &H80000005&
      Caption         =   "Destination Folder"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblNOW 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Check every           minutes."
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label lblTimesChecked 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Number of Times Checked:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Location of File"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "frmCheckforFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public dteTime As Date

Private Sub Check1_Click()

End Sub

Private Sub chkUpdateFile_Click()
Dim chk As Boolean
    lblUpdateWith.Visible = Not lblUpdateWith.Visible
    txtUpdateWith.Visible = Not txtUpdateWith.Visible
    lblUpdateto.Visible = Not txtUpdateTo.Visible
    txtUpdateTo.Visible = Not txtUpdateTo.Visible
    

End Sub

Private Sub cmdBeginCheck_Click()
Timer1.Enabled = Not Timer1.Enabled
dteTime = Now
Call Timer1_Timer
If cmdBeginCheck.Caption = "Begin Checking" Then
    cmdBeginCheck.Caption = "Stop Checking"
Else
    cmdBeginCheck.Caption = "Begin Checking"
End If
End Sub


Private Sub Timer1_Timer()
On Error GoTo BeginCheckError
If cmdBeginCheck.Caption = "Begin Checking" Then Exit Sub
If DateDiff("n", dteTime, Now) > txtNumberofMinutes Then
    lblNOW.Caption = "Last Checked: " & Now
    dteTime = Now
        'check for properties of file
        '(a quick way to find out if there is a file)
        'if there is no file it fires an error which starts the copy process
    VBA.FileSystem.FileDateTime (txtFile)
    txtNumberChecked = txtNumberChecked + 1
End If
Exit Sub
BeginCheckError:
    If Err.Number = 53 Then
        If Me.chkUpdateFile = 1 Then
            COPYFILE
        Else
            Me.SetFocus
            Beep
            MsgBox "Your file is now missing"
        End If
    Else
        MsgBox "Error: " & Err.Description & " " & Err.Number
    End If
End Sub

Private Sub txtFile_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtFile = Data.Files.Item(1)
End Sub

Private Sub txtUpdateTo_gotfocus()
txtUpdateTo = ""
End Sub

Private Sub txtUpdateTo_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtUpdateTo = Data.Files.Item(1)
End Sub

Private Sub txtUpdateWith_GotFocus()
Dim MyString As String, strScrap As String, Counter As Long
txtUpdateWith = ""
    'Update txtUpdateTo with the folder location of file in txtFile
    'by parsing the leftmost of the textbox (getting rid of actual file at end)
    If IsNull(Me.txtFile) = False Then
        If txtFile <> "" Then
            MyString = txtFile
            Do Until Left(strScrap, 1) = "\"
                Counter = Counter + 1
                strScrap = Right(MyString, Counter)
            Loop
            txtUpdateTo = Left(Me.txtFile, Len(txtFile) - Len(strScrap) + 1)
        End If
    End If
End Sub
Private Sub COPYFILE()
Dim MyString As String, strScrap As String, Counter As Integer
strScrap = Me.txtUpdateWith
Do Until Left(strScrap, 1) = "\"
    Counter = Counter + 1
    strScrap = Right(Me.txtUpdateWith, Counter)
Loop
strScrap = Right(strScrap, Len(strScrap) - 1)
FileCopy Me.txtUpdateWith, Me.txtUpdateTo & strScrap

MsgBox "File Copied"
cmdBeginCheck.Caption = "Begin Checking"
End Sub

Private Sub txtUpdateWith_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
txtUpdateWith = Data.Files.Item(1)
End Sub
