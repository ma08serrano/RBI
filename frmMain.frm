VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{94A0E92D-43C0-494E-AC29-FD45948A5221}#1.0#0"; "wiaaut.dll"
Begin VB.Form frmMain 
   Caption         =   "Information Maintenance"
   ClientHeight    =   9825
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   13200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   93
      Top             =   9570
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12330
            Picture         =   "frmMain.frx":08CA
            Text            =   "Ready"
            TextSave        =   "Ready"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   2
            Enabled         =   0   'False
            Picture         =   "frmMain.frx":0AE0
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Picture         =   "frmMain.frx":0CC2
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Picture         =   "frmMain.frx":0F0C
            TextSave        =   "5/10/2012"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Picture         =   "frmMain.frx":114F
            TextSave        =   "11:06 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   78
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1429
      ButtonWidth     =   1455
      ButtonHeight    =   1376
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageButtons"
      DisabledImageList=   "ImageButtons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Edit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Cancel"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clearance"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Report"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logout"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            ImageIndex      =   10
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageButtons 
         Left            =   10800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   10
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1366
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1AAA
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D95
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2165
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25A5
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":288A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F3F
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":32D2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   12855
      Begin VB.Frame Frame13 
         Height          =   2175
         Left            =   10560
         TabIndex        =   94
         Top             =   5160
         Width           =   2175
         Begin VB.Frame Frame14 
            Height          =   1575
            Left            =   120
            TabIndex        =   95
            Top             =   480
            Width           =   1935
            Begin VB.OptionButton optMakeShift 
               Caption         =   "Make S&hift"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   32
               Top             =   1080
               Width           =   1335
            End
            Begin VB.OptionButton optSemiConcrete 
               Caption         =   "S&emi Concrete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   31
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton optConcrete 
               Caption         =   "&Concrete"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   30
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type of House:"
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
            Left            =   240
            TabIndex        =   96
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame16 
         Height          =   1575
         Left            =   120
         TabIndex        =   82
         Top             =   7320
         Width           =   12615
         Begin VB.Frame Frame18 
            Height          =   975
            Left            =   6360
            TabIndex        =   87
            Top             =   480
            Width           =   6135
            Begin VB.TextBox txtMMiddleName 
               Enabled         =   0   'False
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
               Left            =   4200
               TabIndex        =   38
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtMFirstName 
               Enabled         =   0   'False
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
               Left            =   2160
               TabIndex        =   37
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtMSurname 
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   36
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label34 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name:"
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
               Left            =   4200
               TabIndex        =   90
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label Label33 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "First Name:"
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
               Left            =   2160
               TabIndex        =   89
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label32 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Surname:"
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
               TabIndex        =   88
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Frame Frame17 
            Height          =   975
            Left            =   120
            TabIndex        =   83
            Top             =   480
            Width           =   6135
            Begin VB.TextBox txtFMiddleName 
               Enabled         =   0   'False
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
               Left            =   4200
               TabIndex        =   35
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtFFirstName 
               Enabled         =   0   'False
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
               Left            =   2160
               TabIndex        =   34
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtFSurname 
               Enabled         =   0   'False
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
               Left            =   120
               TabIndex        =   33
               Top             =   480
               Width           =   1815
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name:"
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
               Left            =   4200
               TabIndex        =   86
               Top             =   240
               Width           =   1170
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "First Name:"
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
               Left            =   2160
               TabIndex        =   85
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Surname:"
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
               TabIndex        =   84
               Top             =   240
               Width           =   810
            End
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mother's Name"
            Height          =   240
            Left            =   8520
            TabIndex        =   92
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Father's Name"
            Height          =   240
            Left            =   2280
            TabIndex        =   91
            Top             =   240
            Width           =   1530
         End
      End
      Begin VB.Frame Frame11 
         Height          =   975
         Left            =   120
         TabIndex        =   71
         Top             =   720
         Width           =   10335
         Begin VB.TextBox txtRelationship 
            Enabled         =   0   'False
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
            Left            =   7560
            TabIndex        =   4
            Top             =   360
            Width           =   2655
         End
         Begin VB.Frame Frame12 
            Height          =   615
            Left            =   2040
            TabIndex        =   73
            Top             =   240
            Width           =   2415
            Begin VB.OptionButton optYes1 
               Caption         =   "&Yes"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   2
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optNo1 
               Caption         =   "&No"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   3
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblSeparator1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               Height          =   240
               Left            =   1200
               TabIndex        =   81
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Relationship of the Hhead:"
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
            Left            =   5040
            TabIndex        =   74
            Top             =   480
            Width           =   2295
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Head of the Family:"
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
            TabIndex        =   72
            Top             =   480
            Width           =   1665
         End
      End
      Begin VB.Frame Frame9 
         Height          =   1935
         Left            =   10560
         TabIndex        =   66
         Top             =   3240
         Width           =   2175
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "&Browse"
            Enabled         =   0   'False
            Height          =   735
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1935
         End
         Begin VB.CommandButton cmdCapture 
            Caption         =   "&Capture"
            Enabled         =   0   'False
            Height          =   735
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Height          =   2535
         Left            =   10560
         TabIndex        =   65
         Top             =   720
         Width           =   2175
         Begin VB.Image PictureRes 
            BorderStyle     =   1  'Fixed Single
            Height          =   2175
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.TextBox txtResident 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11400
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Height          =   3495
         Left            =   6360
         TabIndex        =   59
         Top             =   3840
         Width           =   4095
         Begin VB.Frame Frame15 
            Height          =   615
            Left            =   1440
            TabIndex        =   77
            Top             =   2760
            Width           =   2415
            Begin VB.OptionButton optNo3 
               Caption         =   "&No"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   29
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optYes3 
               Caption         =   "&Yes"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   28
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblSeparator3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               Height          =   240
               Left            =   1200
               TabIndex        =   80
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.ComboBox cboBloodType 
            Enabled         =   0   'False
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
            ItemData        =   "frmMain.frx":36B0
            Left            =   1440
            List            =   "frmMain.frx":36C0
            TabIndex        =   25
            Top             =   1800
            Width           =   615
         End
         Begin VB.Frame Frame10 
            Height          =   615
            Left            =   1440
            TabIndex        =   68
            Top             =   2160
            Width           =   2415
            Begin VB.OptionButton optNo2 
               Caption         =   "&No"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   1560
               TabIndex        =   27
               Top             =   240
               Width           =   735
            End
            Begin VB.OptionButton optYes2 
               Caption         =   "&Yes"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   240
               TabIndex        =   26
               Top             =   240
               Width           =   735
            End
            Begin VB.Label lblSeparator2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "/"
               Height          =   240
               Left            =   1200
               TabIndex        =   79
               Top             =   240
               Width           =   90
            End
         End
         Begin VB.TextBox txtOccupation 
            Enabled         =   0   'False
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
            Left            =   1440
            TabIndex        =   24
            Top             =   1320
            Width           =   2535
         End
         Begin VB.TextBox txtCitizenship 
            Enabled         =   0   'False
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
            Left            =   1440
            TabIndex        =   23
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox txtZipCode 
            Enabled         =   0   'False
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
            Left            =   1440
            TabIndex        =   22
            Text            =   "2400"
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registered Voter?"
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
            Left            =   120
            TabIndex        =   76
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Blood Type:"
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
            TabIndex        =   70
            Top             =   1920
            Width           =   1035
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PWDs:"
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
            TabIndex        =   67
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Occupation:"
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
            TabIndex        =   62
            Top             =   1440
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Citizenship:"
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
            TabIndex        =   61
            Top             =   960
            Width           =   990
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Zip Code:"
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
            TabIndex        =   60
            Top             =   480
            Width           =   840
         End
      End
      Begin VB.Frame Frame4 
         Height          =   3495
         Left            =   120
         TabIndex        =   53
         Top             =   3840
         Width           =   6135
         Begin MSComCtl2.DTPicker cboBirthdate 
            Height          =   375
            Left            =   1200
            TabIndex        =   14
            Top             =   1320
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   48758785
            CurrentDate     =   41026
         End
         Begin VB.ComboBox cboAttainment 
            Enabled         =   0   'False
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
            ItemData        =   "frmMain.frx":36D1
            Left            =   4320
            List            =   "frmMain.frx":36E7
            TabIndex        =   21
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtAge 
            Enabled         =   0   'False
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
            Left            =   1200
            TabIndex        =   15
            Top             =   1800
            Width           =   615
         End
         Begin VB.Frame Frame6 
            Height          =   1695
            Left            =   4320
            TabIndex        =   58
            Top             =   240
            Width           =   1695
            Begin VB.OptionButton optSeparated 
               Caption         =   "Se&parated"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1320
               Width           =   1455
            End
            Begin VB.OptionButton optWidow 
               Caption         =   "&Widow/er"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   960
               Width           =   1335
            End
            Begin VB.OptionButton optMarried 
               Caption         =   "Ma&rried"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optSingle 
               Caption         =   "&Single"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   17
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.TextBox txtBirthplace 
            Enabled         =   0   'False
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
            Left            =   1200
            TabIndex        =   16
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Frame Frame5 
            Height          =   975
            Left            =   1200
            TabIndex        =   55
            Top             =   240
            Width           =   1455
            Begin VB.OptionButton optFemale 
               Caption         =   "&Female"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   1215
            End
            Begin VB.OptionButton optMale 
               Caption         =   "&Male"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   855
            End
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Highest Educational Attainment:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2400
            TabIndex        =   75
            Top             =   2760
            Width           =   1755
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Age:"
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
            TabIndex        =   69
            Top             =   1920
            Width           =   405
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Civil Status:"
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
            Left            =   3120
            TabIndex        =   64
            Top             =   480
            Width           =   1035
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birthplace:"
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
            TabIndex        =   57
            Top             =   2400
            Width           =   930
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Birthdate:"
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
            TabIndex        =   56
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Gender:"
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
            TabIndex        =   54
            Top             =   480
            Width           =   690
         End
      End
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   47
         Top             =   2760
         Width           =   10335
         Begin VB.TextBox txtProvince 
            Enabled         =   0   'False
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
            Left            =   8400
            TabIndex        =   11
            Text            =   "Pangasinan"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtCity 
            Enabled         =   0   'False
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
            Left            =   6480
            TabIndex        =   10
            Text            =   "Dagupan City"
            Top             =   600
            Width           =   1815
         End
         Begin VB.ComboBox cboBarangay 
            Enabled         =   0   'False
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
            ItemData        =   "frmMain.frx":373F
            Left            =   4320
            List            =   "frmMain.frx":37A0
            TabIndex        =   9
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtStreet 
            Enabled         =   0   'False
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
            Left            =   2040
            TabIndex        =   8
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Residential Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   480
            Width           =   1695
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Province:"
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
            Left            =   8400
            TabIndex        =   51
            Top             =   360
            Width           =   825
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City/Municipality:"
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
            Left            =   6480
            TabIndex        =   50
            Top             =   360
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "Barangay:"
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
            Left            =   4320
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Street/Zone:"
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
            Left            =   2040
            TabIndex        =   48
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   10335
         Begin VB.TextBox txtMiddleName 
            Enabled         =   0   'False
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
            Left            =   7560
            TabIndex        =   7
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtFirstName 
            Enabled         =   0   'False
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
            Left            =   4800
            TabIndex        =   6
            Top             =   600
            Width           =   2655
         End
         Begin VB.TextBox txtSurname 
            Enabled         =   0   'False
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
            TabIndex        =   5
            Top             =   600
            Width           =   2655
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Middle Name:"
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
            Left            =   7560
            TabIndex        =   45
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "First Name:"
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
            Left            =   4800
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Surname:"
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
            Left            =   2040
            TabIndex        =   43
            Top             =   360
            Width           =   810
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resident's Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   1440
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Resident Number"
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9360
         TabIndex        =   63
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   2130
      End
   End
   Begin WIACtl.CommonDialog CommonDialog2 
      Left            =   13200
      Top             =   1320
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu fileUpdateInfo 
         Caption         =   "Update Info"
      End
      Begin VB.Menu filePrinterSettings 
         Caption         =   "Printer &Settings"
         Shortcut        =   ^S
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu fileLogout 
         Caption         =   "&Log Out"
         Shortcut        =   ^L
      End
      Begin VB.Menu fileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu menuView 
      Caption         =   "&View"
      Begin VB.Menu viewClearance 
         Caption         =   "Barangay Clearance"
      End
      Begin VB.Menu viewSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu viewReport 
         Caption         =   "Report"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu helpAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mode, valid As Integer
Dim choice As Integer

Private Sub cmdBrowse_Click()
    On Error GoTo cam_trap
    AddFilter CommonDialog1, "Image Files", "*.gif;*.jpg;*.jpeg;*.png"
    CommonDialog1.ShowOpen
    If Not CommonDialog1.FileName = "" Then
save_again:
        FileCopy CommonDialog1.FileName, App.Path & "\Resident\" & txtResident & ".jpg"
        PictureRes = LoadPicture(App.Path & "\Resident\" & txtResident & ".jpg")
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

Private Sub cmdCapture_Click()
On Error Resume Next
    CommonDialog2.ShowAcquireImage.SaveFile App.Path & "\Resident\" & txtResident & ".jpg"
    PictureRes = LoadPicture(App.Path & "\Resident\" & txtResident & ".jpg")
Exit Sub
com_trap:
    MsgBox Error$, vbExclamation, "Registry of Barangay Inhabitants"
End Sub

Private Sub Form_Load()
    Dim sql As String
    sql = "SELECT COUNT(*) FROM tblRBI"
    Set rs = db.Execute(sql)
    StatusBar1.Panels(1).Text = "Total Record(s): " & rs.Fields(0)
End Sub

Private Sub optYes1_Click()
    If optYes1 = True Then
        txtRelationship = ""
        txtRelationship.Enabled = False
    End If
End Sub

Private Sub optNo1_Click()
    If optNo1 = True Then
        txtRelationship = ""
        txtRelationship.Enabled = True
    End If
        optNo1.Value = True
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sql As String
    valid = 1
    Dim relation, gender, civil, pwds, voter, house As String
    If Button.Caption = "New" Then
        mode = 1
        optYes1.Value = False
        optNo1.Value = False
        txtRelationship = ""
        txtSurname = ""
        txtFirstName = ""
        txtMiddleName = ""
        txtStreet = ""
        cboBarangay = ""
        optMale.Value = False
        optFemale.Value = False
        txtAge = ""
        txtBirthplace = ""
        optSingle.Value = False
        optMarried.Value = False
        optWidow.Value = False
        optSeparated.Value = False
        cboAttainment = ""
        txtCitizenship = ""
        txtOccupation = ""
        cboBloodType = ""
        optYes2.Value = False
        optNo2.Value = False
        optYes3.Value = False
        optNo3.Value = False
        optConcrete.Value = False
        optSemiConcrete.Value = False
        optMakeShift.Value = False
        optSeparated.Value = False
        txtFSurname = ""
        txtFFirstName = ""
        txtFMiddleName = ""
        txtMSurname = ""
        txtMFirstName = ""
        txtMMiddleName = ""
        Toolbar.Buttons(1).Enabled = False
        Toolbar.Buttons(2).Enabled = False
        Toolbar.Buttons(3).Enabled = False
        Toolbar.Buttons(4).Enabled = True
        Toolbar.Buttons(5).Enabled = True
        Toolbar.Buttons(7).Enabled = False
        Toolbar.Buttons(8).Enabled = False
        Toolbar.Buttons(9).Enabled = False
        Toolbar.Buttons(11).Enabled = False
        Toolbar.Buttons(12).Enabled = False
        txtResident.Enabled = False
        optYes1.Enabled = True
        optNo1.Enabled = True
        txtRelationship.Enabled = True
        txtSurname.Enabled = True
        txtFirstName.Enabled = True
        txtMiddleName.Enabled = True
        txtStreet.Enabled = True
        cboBarangay.Enabled = True
        optMale.Enabled = True
        optFemale.Enabled = True
        cboBirthdate.Enabled = True
        txtBirthplace.Enabled = True
        optSingle.Enabled = True
        optMarried.Enabled = True
        optWidow.Enabled = True
        optSeparated.Enabled = True
        cboAttainment.Enabled = True
        txtCitizenship.Enabled = True
        txtOccupation.Enabled = True
        cboBloodType.Enabled = True
        optYes2.Enabled = True
        optNo2.Enabled = True
        optYes3.Enabled = True
        optNo3.Enabled = True
        cmdCapture.Enabled = True
        cmdBrowse.Enabled = True
        optConcrete.Enabled = True
        optSemiConcrete.Enabled = True
        optMakeShift.Enabled = True
        txtFSurname.Enabled = True
        txtFFirstName.Enabled = True
        txtFMiddleName.Enabled = True
        txtMSurname.Enabled = True
        txtMFirstName.Enabled = True
        txtMMiddleName.Enabled = True
        sql = "SELECT * FROM tblRBI ORDER BY Resident DESC"
        Set rs = db.Execute(sql)
        If rs.EOF Then
            txtResident = 1
        Else
            txtResident = rs!Resident + 1
        End If
    ElseIf Button.Caption = "Edit" Then
        If txtResident = "" Then
            MsgBox "Please enter Resident Number.", vbExclamation, "Registry of Barangay Inhabitants"
        Else
            txtResident.Enabled = False
            optYes1.Enabled = True
            optNo1.Enabled = True
            txtRelationship.Enabled = True
            txtSurname.Enabled = True
            txtFirstName.Enabled = True
            txtMiddleName.Enabled = True
            txtStreet.Enabled = True
            cboBarangay.Enabled = True
            optMale.Enabled = True
            optFemale.Enabled = True
            cboBirthdate.Enabled = True
            txtBirthplace.Enabled = True
            optSingle.Enabled = True
            optMarried.Enabled = True
            optWidow.Enabled = True
            optSeparated.Enabled = True
            cboAttainment.Enabled = True
            txtCitizenship.Enabled = True
            txtOccupation.Enabled = True
            cboBloodType.Enabled = True
            optYes2.Enabled = True
            optNo2.Enabled = True
            optYes3.Enabled = True
            optNo3.Enabled = True
            cmdCapture.Enabled = True
            cmdBrowse.Enabled = True
            optConcrete.Enabled = True
            optSemiConcrete.Enabled = True
            optMakeShift.Enabled = True
            txtFSurname.Enabled = True
            txtFFirstName.Enabled = True
            txtFMiddleName.Enabled = True
            txtMSurname.Enabled = True
            txtMFirstName.Enabled = True
            txtMMiddleName.Enabled = True
            Toolbar.Buttons(1).Enabled = False
            Toolbar.Buttons(2).Enabled = False
            Toolbar.Buttons(3).Enabled = False
            Toolbar.Buttons(4).Enabled = True
            Toolbar.Buttons(5).Enabled = True
            Toolbar.Buttons(7).Enabled = False
            Toolbar.Buttons(8).Enabled = False
            Toolbar.Buttons(9).Enabled = False
            Toolbar.Buttons(11).Enabled = False
            Toolbar.Buttons(12).Enabled = False
            mode = 2
        End If
     ElseIf Button.Caption = "Delete" Then
        On Error Resume Next
        If txtResident = "" Then
            MsgBox "Please enter Resident Number first to be able to delete a record.", vbExclamation, "Registry of barangay Inhabitants"
        Else
            sql = "SELECT * FROM tblRBI WHERE Resident=" & txtResident
            Set rs = db.Execute(sql)
                If rs.EOF Then
                    MsgBox "No such Resident Number found, Please try again!", vbExclamation, "Registry of Barangay Inhabitants"
        Else
            choice = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Registry of Barangay Inhabitants")
                If choice = vbYes Then
                    sql = "DELETE * FROM tblRBI WHERE Resident = " & txtResident
                    db.Execute sql
                    MsgBox "Record has been deleted successfully.", vbInformation, "Registry of Barangay Inhabitants"
                    txtResident = ""
                    optYes1.Value = False
                    optNo1.Value = False
                    txtRelationship = ""
                    txtSurname = ""
                    txtFirstName = ""
                    txtMiddleName = ""
                    txtStreet = ""
                    cboBarangay = ""
                    optMale.Value = False
                    optFemale.Value = False
                    txtAge = ""
                    txtBirthplace = ""
                    optSingle.Value = False
                    optMarried.Value = False
                    optWidow.Value = False
                    optSeparated.Value = False
                    cboAttainment = ""
                    txtCitizenship = ""
                    txtOccupation = ""
                    cboBloodType = ""
                    optYes2.Value = False
                    optNo2.Value = False
                    optYes3.Value = False
                    optNo3.Value = False
                    optConcrete.Value = False
                    optSemiConcrete.Value = False
                    optMakeShift.Value = False
                    txtFSurname = ""
                    txtFFirstName = ""
                    txtFMiddleName = ""
                    txtMSurname = ""
                    txtMFirstName = ""
                    txtMMiddleName = ""
                    Toolbar.Buttons(1).Enabled = True
                    Toolbar.Buttons(2).Enabled = False
                    Toolbar.Buttons(3).Enabled = True
                    Toolbar.Buttons(4).Enabled = False
                    Toolbar.Buttons(5).Enabled = False
                    Toolbar.Buttons(7).Enabled = True
                    Toolbar.Buttons(8).Enabled = True
                    Toolbar.Buttons(9).Enabled = True
                    Toolbar.Buttons(11).Enabled = True
                    Toolbar.Buttons(12).Enabled = True
                End If
                End If
        End If
     ElseIf Button.Caption = "Save" Then
        If optNo1 = True Then
            If txtRelationship = "" Then
                valid = 0
            End If
        End If
            If txtSurname = "" Then
                valid = 0
            ElseIf txtFirstName = "" Then
                valid = 0
            ElseIf txtMiddleName = "" Then
                valid = 0
            ElseIf txtStreet = "" Then
                valid = 0
            ElseIf cboBarangay = "" Then
                valid = 0
            ElseIf txtBirthplace = "" Then
                valid = 0
            ElseIf cboAttainment = "" Then
                valid = 0
            ElseIf txtCitizenship = "" Then
                valid = 0
            ElseIf txtOccupation = "" Then
                valid = 0
            ElseIf cboBloodType = "" Then
                valid = 0
            ElseIf txtFSurname = "" Then
                valid = 0
            ElseIf txtFFirstName = "" Then
                valid = 0
            ElseIf txtFMiddleName = "" Then
                valid = 0
            ElseIf txtMSurname = "" Then
                valid = 0
            ElseIf txtMFirstName = "" Then
                valid = 0
            ElseIf txtMMiddleName = "" Then
                valid = 0
            End If
        If mode = 1 Then
            If valid = 0 Then
                MsgBox "Please fill up all the required fields.", vbExclamation, "Registry of Barangay Inhabitants"
            Else
                If optNo1 = True Then
                    relation = txtRelationship
                Else
                    relation = ""
                End If
                If optMale = True Then
                    gender = "Male"
                Else
                    gender = "Female"
                End If
                If optSingle = True Then
                    civil = "Single"
                ElseIf optMarried = True Then
                    civil = "Married"
                ElseIf optWidow = True Then
                    civil = "Widow/er"
                ElseIf optSeparated = True Then
                    civil = "Separated"
                End If
                If optNo2 = True Then
                    pwds = "No"
                Else
                    pwds = "Yes"
                End If
                If optNo3 = True Then
                    voter = "No"
                Else
                    voter = "Yes"
                End If
                If optConcrete = True Then
                    house = "Concrete"
                ElseIf optSemiConcrete = True Then
                    house = "Semi Concrete"
                ElseIf optMakeShift = True Then
                    house = "Make Shift"
                End If
                sql = "INSERT INTO tblRBI VALUES(" & txtResident & ",'" & relation & "','" & txtSurname & "','" & txtFirstName & "','" & txtMiddleName & "','" & txtStreet & "','" & cboBarangay & "','" & txtCity & "','" & txtProvince & "','" & gender & "','" & cboBirthdate & "','" & txtBirthplace & "','" & civil & "','" & cboAttainment & "','" & txtZipCode & "','" & txtCitizenship & "','" & txtOccupation & "','" & cboBloodType & "','" & pwds & "','" & voter & "','" & house & "','" & txtFSurname & "','" & txtFFirstName & "','" & txtFMiddleName & "','" & txtMSurname & "','" & txtMFirstName & "','" & txtMMiddleName & "')"
                db.Execute sql
                MsgBox "Records has been saved.", vbInformation, "Registry of Barangay Inhabitants"
                Toolbar.Buttons(1).Enabled = True
                Toolbar.Buttons(2).Enabled = False
                Toolbar.Buttons(3).Enabled = True
                Toolbar.Buttons(4).Enabled = False
                Toolbar.Buttons(5).Enabled = False
                Toolbar.Buttons(7).Enabled = True
                Toolbar.Buttons(8).Enabled = True
                Toolbar.Buttons(9).Enabled = True
                Toolbar.Buttons(11).Enabled = True
                Toolbar.Buttons(12).Enabled = True
                txtResident = ""
                optYes1.Value = False
                optNo1.Value = False
                txtRelationship = ""
                txtSurname = ""
                txtFirstName = ""
                txtMiddleName = ""
                txtStreet = ""
                cboBarangay = ""
                optMale.Value = False
                optFemale.Value = False
                txtAge = ""
                txtBirthplace = ""
                optSingle.Value = False
                optMarried.Value = False
                optWidow.Value = False
                optSeparated.Value = False
                cboAttainment = ""
                txtCitizenship = ""
                txtOccupation = ""
                cboBloodType = ""
                optYes2.Value = False
                optNo2.Value = False
                optYes3.Value = False
                optNo3.Value = False
                optConcrete.Value = False
                optSemiConcrete.Value = False
                optMakeShift.Value = False
                txtFSurname = ""
                txtFFirstName = ""
                txtFMiddleName = ""
                txtMSurname = ""
                txtMFirstName = ""
                txtMMiddleName = ""
                txtResident.Enabled = True
                optYes1.Enabled = False
                optNo1.Enabled = False
                txtRelationship.Enabled = False
                txtSurname.Enabled = False
                txtFirstName.Enabled = False
                txtMiddleName.Enabled = False
                txtStreet.Enabled = False
                cboBarangay.Enabled = False
                optMale.Enabled = False
                optFemale.Enabled = False
                cboBirthdate.Enabled = False
                txtBirthplace.Enabled = False
                optSingle.Enabled = False
                optMarried.Enabled = False
                optWidow.Enabled = False
                optSeparated.Enabled = False
                cboAttainment.Enabled = False
                txtCitizenship.Enabled = False
                txtOccupation.Enabled = False
                cboBloodType.Enabled = False
                optYes2.Enabled = False
                optNo2.Enabled = False
                optYes3.Enabled = False
                optNo3.Enabled = False
                cmdCapture.Enabled = False
                cmdBrowse.Enabled = False
                optConcrete.Enabled = False
                optSemiConcrete.Enabled = False
                optMakeShift.Enabled = False
                txtFSurname.Enabled = False
                txtFFirstName.Enabled = False
                txtFMiddleName.Enabled = False
                txtMSurname.Enabled = False
                txtMFirstName.Enabled = False
                txtMMiddleName.Enabled = False
            End If
            ElseIf mode = 2 Then
                If valid = 0 Then
                    MsgBox "Please fill up all the required fields.", vbExclamation, "Registry of Barangay Inhabitants"
                Else
                    If optNo1 = True Then
                        relation = txtRelationship
                    Else
                        relation = ""
                    End If
                    If optMale = True Then
                        gender = "Male"
                    Else
                        gender = "Female"
                    End If
                    If optSingle = True Then
                        civil = "Single"
                    ElseIf optMarried = True Then
                        civil = "Married"
                    ElseIf optWidow = True Then
                        civil = "Widow/er"
                    ElseIf optSeparated = True Then
                        civil = "Separated"
                    End If
                    If optNo2 = True Then
                        pwds = "No"
                    Else
                        pwds = "Yes"
                    End If
                    If optNo3 = True Then
                        voter = "No"
                    Else
                        voter = "Yes"
                    End If
                    If optConcrete = True Then
                        house = "Concrete"
                    ElseIf optSemiConcrete = True Then
                        house = "Semi Concrete"
                    ElseIf optMakeShift = True Then
                        house = "Make Shift"
                    End If
                sql = "UPDATE tblRBI SET Relationship='" & txtRelationship & "', Surname='" & txtSurname & "', FirstName= '" & txtFirstName & "', MiddleName = '" & txtMiddleName & "', Street='" & txtStreet & "', Barangay='" & cboBarangay & "', gender='" & gender & "', Birthdate='" & cboBirthdate & "', BirthPlace='" & txtBirthplace & "', civil='" & civil & "', Attainment='" & cboAttainment & "', Citizenship='" & txtCitizenship & "', Occupation='" & txtOccupation & "', BloodType='" & cboBloodType & "', PWDs='" & pwds & "', Voter='" & voter & "', House='" & house & "', FSurname='" & txtFSurname & "', FFirstName= '" & txtFFirstName & "', FMiddleName = '" & txtFMiddleName & "', MSurname='" & txtMSurname & "', MFirstName= '" & txtMFirstName & "', MMiddleName = '" & txtMMiddleName & "' WHERE Resident =" & txtResident
                db.Execute sql
                MsgBox "Resident Info has been updated successfully.", vbInformation, "Registry of Barangay Inhabitants"
                Toolbar.Buttons(1).Enabled = True
                Toolbar.Buttons(2).Enabled = False
                Toolbar.Buttons(3).Enabled = True
                Toolbar.Buttons(4).Enabled = False
                Toolbar.Buttons(5).Enabled = False
                Toolbar.Buttons(7).Enabled = True
                Toolbar.Buttons(8).Enabled = True
                Toolbar.Buttons(9).Enabled = True
                Toolbar.Buttons(11).Enabled = True
                Toolbar.Buttons(12).Enabled = True
                txtResident = ""
                optYes1.Value = False
                optNo1.Value = False
                txtRelationship = ""
                txtSurname = ""
                txtFirstName = ""
                txtMiddleName = ""
                txtStreet = ""
                cboBarangay = ""
                optMale.Value = False
                optFemale.Value = False
                txtAge = ""
                txtBirthplace = ""
                optSingle.Value = False
                optMarried.Value = False
                optWidow.Value = False
                optSeparated.Value = False
                cboAttainment = ""
                txtCitizenship = ""
                txtOccupation = ""
                cboBloodType = ""
                optYes2.Value = False
                optNo2.Value = False
                optYes3.Value = False
                optNo3.Value = False
                optConcrete.Value = False
                optSemiConcrete.Value = False
                optMakeShift.Value = False
                txtFSurname = ""
                txtFFirstName = ""
                txtFMiddleName = ""
                txtMSurname = ""
                txtMFirstName = ""
                txtMMiddleName = ""
                txtResident.Enabled = True
                optYes1.Enabled = False
                optNo1.Enabled = False
                txtRelationship.Enabled = False
                txtSurname.Enabled = False
                txtFirstName.Enabled = False
                txtMiddleName.Enabled = False
                txtStreet.Enabled = False
                cboBarangay.Enabled = False
                optMale.Enabled = False
                optFemale.Enabled = False
                cboBirthdate.Enabled = False
                txtBirthplace.Enabled = False
                optSingle.Enabled = False
                optMarried.Enabled = False
                optWidow.Enabled = False
                optSeparated.Enabled = False
                cboAttainment.Enabled = False
                txtCitizenship.Enabled = False
                txtOccupation.Enabled = False
                cboBloodType.Enabled = False
                optYes2.Enabled = False
                optNo2.Enabled = False
                optYes3.Enabled = False
                optNo3.Enabled = False
                cmdCapture.Enabled = False
                cmdBrowse.Enabled = False
                optConcrete.Enabled = False
                optSemiConcrete.Enabled = False
                optMakeShift.Enabled = False
                txtFSurname.Enabled = False
                txtFFirstName.Enabled = False
                txtFMiddleName.Enabled = False
                txtMSurname.Enabled = False
                txtMFirstName.Enabled = False
                txtMMiddleName.Enabled = False
                End If
        End If
     ElseIf Button.Caption = "Cancel" Then
        Toolbar.Buttons(1).Enabled = True
        Toolbar.Buttons(2).Enabled = False
        Toolbar.Buttons(3).Enabled = True
        Toolbar.Buttons(4).Enabled = False
        Toolbar.Buttons(5).Enabled = False
        Toolbar.Buttons(7).Enabled = True
        Toolbar.Buttons(8).Enabled = True
        Toolbar.Buttons(9).Enabled = True
        Toolbar.Buttons(11).Enabled = True
        Toolbar.Buttons(12).Enabled = True
        txtResident.Enabled = True
        optYes1.Enabled = False
        optNo1.Enabled = False
        txtRelationship.Enabled = False
        txtSurname.Enabled = False
        txtFirstName.Enabled = False
        txtMiddleName.Enabled = False
        txtStreet.Enabled = False
        cboBarangay.Enabled = False
        optMale.Enabled = False
        optFemale.Enabled = False
        cboBirthdate.Enabled = False
        txtBirthplace.Enabled = False
        optSingle.Enabled = False
        optMarried.Enabled = False
        optWidow.Enabled = False
        optSeparated.Enabled = False
        cboAttainment.Enabled = False
        txtCitizenship.Enabled = False
        txtOccupation.Enabled = False
        cboBloodType.Enabled = False
        optYes2.Enabled = False
        optNo2.Enabled = False
        optYes3.Enabled = False
        optNo3.Enabled = False
        cmdCapture.Enabled = False
        cmdBrowse.Enabled = False
        optConcrete.Enabled = False
        optSemiConcrete.Enabled = False
        optMakeShift.Enabled = False
        txtFSurname.Enabled = False
        txtFFirstName.Enabled = False
        txtFMiddleName.Enabled = False
        txtMSurname.Enabled = False
        txtMFirstName.Enabled = False
        txtMMiddleName.Enabled = False
        txtResident = ""
        optYes1.Value = False
        optNo1.Value = False
        txtRelationship = ""
        txtSurname = ""
        txtFirstName = ""
        txtMiddleName = ""
        txtStreet = ""
        cboBarangay = ""
        optMale.Value = False
        optFemale.Value = False
        txtAge = ""
        txtBirthplace = ""
        optSingle.Value = False
        optMarried.Value = False
        optWidow.Value = False
        optSeparated.Value = False
        cboAttainment = ""
        txtCitizenship = ""
        txtOccupation = ""
        cboBloodType = ""
        optYes2.Value = False
        optNo2.Value = False
        optYes3 = False
        optNo3.Value = False
        optConcrete.Value = False
        optSemiConcrete.Value = False
        optMakeShift.Value = False
        txtFSurname = ""
        txtFFirstName = ""
        txtFMiddleName = ""
        txtMSurname = ""
        txtMFirstName = ""
        txtMMiddleName = ""
    ElseIf Button.Caption = "Search" Then
        Load frmSearch
        frmSearch.Show vbModal
    ElseIf Button.Caption = "Report" Then
        Load frmReport
        frmReport.Show vbModal
    ElseIf Button.Caption = "Clearance" Then
        Load frmClearance
        frmClearance.Show vbModal
    ElseIf Button.Caption = "Logout" Then
        choice = MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo, "Registry of Barangay Inhabitants")
            If choice = vbYes Then
                Me.Hide
                Unload Me
                Load frmLogin
                frmLogin.Show
            End If
    Else
         choice = MsgBox("Are you sure you want to terminate this system?", vbQuestion + vbYesNo, "Registry of Barangay Inhabitants")
            If choice = vbYes Then
                End
            End If
    End If
End Sub

Private Sub txtResident_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 13 Then
        Dim sql As String
        sql = "SELECT * FROM tblRBI WHERE Resident=" & txtResident
        Set rs = db.Execute(sql)
        Toolbar.Buttons(1).Enabled = True
        Toolbar.Buttons(2).Enabled = True
        Toolbar.Buttons(3).Enabled = True
        Toolbar.Buttons(4).Enabled = False
        Toolbar.Buttons(5).Enabled = False
        Toolbar.Buttons(7).Enabled = True
        Toolbar.Buttons(8).Enabled = True
        Toolbar.Buttons(9).Enabled = True
        Toolbar.Buttons(11).Enabled = True
        Toolbar.Buttons(12).Enabled = True
        If rs.EOF Then
            MsgBox "Resident Number not found.", vbExclamation, "Registry of Barangay Inhabitants"
            txtResident = ""
            optYes1.Value = False
            optNo1.Value = False
            txtRelationship = ""
            txtSurname = ""
            txtFirstName = ""
            txtMiddleName = ""
            txtStreet = ""
            cboBarangay = ""
            optMale.Value = False
            optFemale.Value = False
            txtAge = ""
            txtBirthplace = ""
            optSingle.Value = False
            optMarried.Value = False
            optWidow.Value = False
            optSeparated.Value = False
            cboAttainment = ""
            txtCitizenship = ""
            txtOccupation = ""
            cboBloodType = ""
            optYes2.Value = False
            optNo2.Value = False
            optYes3.Value = False
            optNo3.Value = False
            optConcrete.Value = False
            optSemiConcrete.Value = False
            optMakeShift.Value = False
            txtFSurname = ""
            txtFFirstName = ""
            txtFMiddleName = ""
            txtMSurname = ""
            txtMFirstName = ""
            txtMMiddleName = ""
        Else
            txtSurname = rs!Surname
            txtFirstName = rs!FirstName
            txtMiddleName = rs!MiddleName
            txtStreet = rs!Street
            cboBarangay = rs!Barangay
            cboBirthdate = rs!Birthdate
            txtBirthplace = rs!Birthplace
            cboAttainment = rs!Attainment
            txtCitizenship = rs!Citizenship
            txtOccupation = rs!Occupation
            cboBloodType = rs!BloodType
            txtFSurname = rs!FSurname
            txtFFirstName = rs!FFirstName
            txtFMiddleName = rs!FMiddleName
            txtMSurname = rs!MSurname
            txtMFirstName = rs!MFirstName
            txtMMiddleName = rs!MMiddleName
            If rs!Relationship = "" Then
                optYes1 = True
                txtRelationship = ""
            Else
                optNo1 = True
                txtRelationship = rs!Relationship
                txtRelationship.Enabled = False
            End If
            If rs!gender = "Male" Then
                optMale = True
            Else
                optFemale = True
            End If
            If rs!civil = "Single" Then
                optSingle = True
            ElseIf rs!civil = "Married" Then
                optMarried = True
            ElseIf rs!civil = "Widow/er" Then
                optWidow = True
            ElseIf rs!civil = "Separated" Then
                optSeparated = True
            End If
            If rs!pwds = "No" Then
                optNo2 = True
            Else
                optYes2 = True
            End If
            If rs!voter = "No" Then
                optNo3 = True
            Else
                optYes3 = True
            End If
            If rs!house = "Concrete" Then
                optConcrete = True
            ElseIf rs!house = "Semi Concrete" Then
                optSemiConcrete = True
            ElseIf rs!house = "Make Shift" Then
                optMakeShift = True
            End If
            txtAge = DateDiff("yyyy", rs!Birthdate, Now)
            If Dir(App.Path & "\Resident\" & txtResident & ".jpg") <> "" Then
                PictureRes = LoadPicture(App.Path & "\Resident\" & txtResident & ".jpg")
            Else
                PictureRes = LoadPicture(App.Path & "\Resident\no_person.jpg")
            End If
        End If
    End If
End Sub

Private Sub fileUpdateInfo_Click()
    Load frmUpdateInfo
    frmUpdateInfo.Show vbModal
End Sub

Private Sub filePrinterSettings_Click()
    CommonDialog1.ShowPrinter
End Sub

Private Sub fileLogout_Click()
    choice = MsgBox("Are you sure you want to log out?", vbQuestion + vbYesNo, "Registry of Barangay Inhabitants")
        If choice = vbYes Then
            Load frmLogin
            frmLogin.Show
            Me.Hide
            Unload Me
        End If
End Sub

Private Sub fileExit_Click()
    choice = MsgBox("Are you sure you want to terminate this system?", vbQuestion + vbYesNo, "Registry of Barangay Inhabitants")
        If choice = vbYes Then
            End
        End If
End Sub

Private Sub viewClearance_Click()
    Load frmClearance
    frmClearance.Show vbModal
End Sub

Private Sub viewSearch_Click()
    Load frmSearch
    frmSearch.Show vbModal
End Sub

Private Sub viewReport_Click()
    Load frmReport
    frmReport.Show vbModal
End Sub

Private Sub helpAbout_Click()
    Load frmAbout
    frmAbout.Show vbModal
End Sub
