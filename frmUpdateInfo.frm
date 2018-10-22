VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmUpdateInfo 
   Caption         =   "Update Info"
   ClientHeight    =   3450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9015
   Icon            =   "frmUpdateInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1815
      Left            =   5280
      TabIndex        =   12
      Top             =   720
      Width           =   3615
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "&Browse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Barangay Logo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1350
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1455
         Left            =   1680
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   8775
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5400
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5055
      Begin VB.TextBox txtSecretary 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtCaptain 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   840
         Width           =   2895
      End
      Begin VB.ComboBox cboBarangay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmUpdateInfo.frx":08CA
         Left            =   2040
         List            =   "frmUpdateInfo.frx":092B
         TabIndex        =   1
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Barangay Secretary:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1740
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Barangay Captain:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Barangay Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   1410
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE SYSTEM INFORMATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   10
      Top             =   240
      Width           =   3555
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   240
      Picture         =   "frmUpdateInfo.frx":0A94
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   720
      Left            =   6720
      Picture         =   "frmUpdateInfo.frx":0E0F
      Top             =   0
      Width           =   2310
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmUpdateInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
   On Error GoTo cam_trap
    AddFilter CommonDialog1, "Image Files", "*.gif;*.jpg;*.jpeg;*.png"
    CommonDialog1.ShowOpen
    If Not CommonDialog1.FileName = "" Then
save_again:
        FileCopy CommonDialog1.FileName, App.Path & "\Resident\brgy_logo.jpg"
        Image1 = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
    End If
    Exit Sub
cam_trap:
        If Err.Number = 76 Then
            MkDir App.Path & "\Resident"
            GoTo save_again
        End If
End Sub
Private Sub AddFilter(ByVal dlg As CommonDialog, ByVal filter_title As String, ByVal filter_value As String)
Dim txt As String

    txt = dlg.Filter
    If Len(txt) > 0 Then txt = txt & "|"
    txt = txt & filter_title & " (" & filter_value & ")|" & _
        filter_value
    dlg.Filter = txt
End Sub

Private Sub cmdSave_Click()
    If txtCaptain = "" Or txtSecretary = "" Then
        MsgBox "Please fill up all necessary information and try again.", vbExclamation, "Registry of Barangay Inahabitants"
    Else
        SaveSetting App.EXEName, "Config", "Barangay Name", cboBarangay.List(cboBarangay.ListIndex)
        SaveSetting App.EXEName, "Config", "Captain", txtCaptain
        SaveSetting App.EXEName, "Config", "Secretary", txtSecretary
        MsgBox "System Configuration has been update successfully.", vbInformation, "Registry of Barangay Inhabitants"
    End If
End Sub

Private Sub Form_Load()
    If GetSetting(App.EXEName, "Config", "Barangay Name") = "" Then
        cboBarangay.ListIndex = 0
    Else
        cboBarangay = GetSetting(App.EXEName, "Config", "Barangay Name")
    End If
    If GetSetting(App.EXEName, "Config", "Captain") = "" Then
        txtCaptain = ""
    Else
        txtCaptain = GetSetting(App.EXEName, "Config", "Captain")
    End If
    If GetSetting(App.EXEName, "Config", "Secretary") = "" Then
        txtSecretary = ""
    Else
        txtSecretary = GetSetting(App.EXEName, "Config", "Secretary")
    End If
    If Dir(App.Path & "\Resident\brgy_logo.jpg") <> "" Then
        Image1 = LoadPicture(App.Path & "\Resident\brgy_logo.jpg")
    Else
        Image1 = LoadPicture(App.Path & "\Resident\no_logo.jpg")
    End If
End Sub
