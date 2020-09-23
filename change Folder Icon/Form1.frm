VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Folder's Icon By Priyan"
   ClientHeight    =   4230
   ClientLeft      =   2085
   ClientTop       =   2385
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7755
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   6600
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   9
      Top             =   2760
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3600
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   8
      Top             =   600
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set Default"
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change Custom Icon"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2160
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Select the folder"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Icon Path "
      Height          =   195
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------
'This is a simple program that changes the icon of any folders
'
'Programmed by priyan
'
'htpp://priyan.netfirms.com
'priyanrrajeevan@rediffmail.com
'----------------------------------------------------------------------------------------------------------------------------------------


Private Sub Command1_Click()
With Me.CommonDialog1
.DialogTitle = "Select an icon"
.Filter = "Icon files,Exe files(*.ico,*.exe)|*.ico;*.exe"
.filename = ""
.ShowOpen
If .filename <> "" Then
    Text1.Text = .filename
    ExtractIcon .filename, Picture2, 32
End If
End With
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
   If Len(Dir1.Path) > 3 Then 'Check selected is not a drive
    WRITEINI ".ShellClassInfo", "iconfile", Text1.Text, Dir1.Path & "\desktop.ini"
    WRITEINI ".ShellClassInfo", "iconindex", "0", Dir1.Path & "\desktop.ini"
    SetAttr Dir1.Path, vbSystem 'This is important
    MsgBox "Icon Changed", vbInformation, App.Title
    ExtractIcon Dir1.Path, Picture1, 32
' You should set the directory to system else your custom icon will not work
    Else
        MsgBox "Select a valied folder", vbCritical, App.Title
        
    End If
Else
    MsgBox "Select a icon", vbCritical, App.Title
    Command1.SetFocus
End If
End Sub

Private Sub Command3_Click()
On Error Resume Next
Kill Dir1.Path & "\desktop.ini"
SetAttr Dir1.Path, vbArchive
ExtractIcon Dir1.Path, Picture1, 32
End Sub

Private Sub Dir1_Change()
Label1.Caption = Dir1.Path
ExtractIcon Dir1.Path, Picture1, 32
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Label1.Caption = ""
End Sub

Private Sub Label4_Click()
frmabout.Show vbModal
End Sub
