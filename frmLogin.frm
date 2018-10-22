VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   2760
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1630.702
   ScaleMode       =   0  'User
   ScaleWidth      =   4408.351
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txtPassword 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1320
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   720
            Width           =   2565
         End
         Begin VB.TextBox txtUsername 
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   1320
            TabIndex        =   1
            Top             =   240
            Width           =   2565
         End
         Begin VB.CheckBox chkRemember 
            Caption         =   "Remember Username"
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
            Left            =   1200
            TabIndex        =   5
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Username:"
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
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Password:"
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
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   840
            Width           =   885
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4215
         Begin VB.CommandButton cmdOK 
            Caption         =   "&Login"
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
            Height          =   390
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   1140
         End
         Begin VB.CommandButton cmdCancel 
            Cancel          =   -1  'True
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
            Height          =   375
            Left            =   2880
            TabIndex        =   4
            Top             =   240
            Width           =   1140
         End
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkRemember_Click()
    If chkRemember = 1 Then
        SaveSetting App.EXEName, "config", "username", txtUsername
    Else
        SaveSetting App.EXEName, "config", "username", ""
    End If
End Sub

Private Sub Form_Load()
    If Not GetSetting(App.EXEName, "config", "username") = "" Then
        txtUsername = GetSetting(App.EXEName, "config", "username")
        chkRemember.Value = 1
    End If
End Sub

Private Sub cmdOK_Click()
    If txtUsername = "" Then
        MsgBox "Please enter your Username.", vbExclamation, "Registry of Barangay Inhabitants"
    ElseIf txtPassword = "" Then
        MsgBox "Please enter your Password.", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        Dim sql As String
        sql = "SELECT * FROM tblLogin WHERE Username = '" & txtUsername & "' AND Password = '" & txtPassword & "'"
        Set rs = db.Execute(sql)
        If rs.EOF Then
            MsgBox "Invalid Username or Password, Please try again!", vbExclamation, "Registry of Barangay Inhabitants"
        Else
            Load frmMain
            frmMain.Show
            Me.Hide
            Unload Me
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    If MsgBox("Are you sure you want to quit?", vbYesNo + vbQuestion, "Registry of Barangay Inhabitants") = vbYes Then
        End
    End If
End Sub
