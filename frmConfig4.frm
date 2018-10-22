VERSION 5.00
Begin VB.Form frmConfig4 
   Caption         =   "System Configuration Wizard"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   Icon            =   "frmConfig4.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Back"
      Enabled         =   0   'False
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
      Left            =   2880
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Finish"
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
      Left            =   4440
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
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
      Left            =   4320
      TabIndex        =   0
      Top             =   2040
      Width           =   2895
   End
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
      Left            =   4320
      TabIndex        =   1
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   120
      Picture         =   "frmConfig4.frx":08CA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registry of Barangay Inhabitants version 1.2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3000
      TabIndex        =   8
      Top             =   240
      Width           =   4140
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfig4.frx":5F4C
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2880
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Barangay Captain:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   2760
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Secretary:"
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
      Left            =   3360
      TabIndex        =   5
      Top             =   2760
      Width           =   885
   End
End
Attribute VB_Name = "frmConfig4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Load frmConfig3
    frmConfig3.Show
    Me.Hide
End Sub

Private Sub Command2_Click()
    If txtCaptain = "" Or txtSecretary = "" Then
        MsgBox "Please fill up all the required fields.", vbExclamation, "Registry of Barangay Inhabitants"
    Else
        bc = txtCaptain
        bs = txtSecretary
        SaveSetting App.EXEName, "Config", "Captain", bc
        SaveSetting App.EXEName, "Config", "Secretary", bs
        SaveSetting App.EXEName, "Config", "Barangay Name", bn
        SaveSetting App.EXEName, "Config", "Configured", "True"
        Dim sql As String
        sql = "INSERT INTO tblLogin VALUES('" & un & "','" & pw & "')"
        db.Execute sql
        MsgBox "System has been configured successfully." & vbNewLine & "In order to take this effect immideately, the system must be restarted.", vbInformation + vbOKOnly, "Registry of Barangay Inhabitants"
        End
    End If
End Sub

Private Sub Command3_Click()
     Dim choice As Integer
    choice = MsgBox("The system has not been configured successfully." & vbNewLine & "Are you sure you want to terminate this?", vbYesNo + vbQuestion, "Registry of Barangay Inhabitants")
    If choice = vbYes Then
        End
    End If
End Sub
