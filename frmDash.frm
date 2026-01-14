VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDash 
   Caption         =   "Dash Board"
   ClientHeight    =   6225
   ClientLeft      =   195
   ClientTop       =   540
   ClientWidth     =   13605
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   8.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmDash.frx":0000
   ScaleHeight     =   6225
   ScaleWidth      =   13605
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   18240
      Top             =   2160
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Duncons\Desktop\Steve\ITAsset.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Duncons\Desktop\Steve\ITAsset.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Perpetua Titling MT"
         Size            =   8.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame frameUser 
      BackColor       =   &H00404000&
      Caption         =   "USERS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   8055
      Left            =   5880
      TabIndex        =   62
      Top             =   2760
      Width           =   14295
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   11400
         Top             =   6960
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   1
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Duncons\Desktop\Steve\ITAsset.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Duncons\Desktop\Steve\ITAsset.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Perpetua Titling MT"
            Size            =   8.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdLogs 
         Caption         =   "LOGS"
         Height          =   735
         Left            =   11400
         TabIndex        =   77
         Top             =   5400
         Width           =   2175
      End
      Begin VB.ComboBox cmbUDepartment 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3240
         TabIndex        =   74
         Text            =   "SELECT"
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txtUUserID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11520
         TabIndex        =   71
         Top             =   2640
         Width           =   2415
      End
      Begin VB.ComboBox cmbURole 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11400
         TabIndex        =   69
         Text            =   "SELECT"
         Top             =   960
         Width           =   2655
      End
      Begin VB.TextBox txtUPassword 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   66
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtUUserName 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3240
         TabIndex        =   64
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label16 
         BackColor       =   &H00404000&
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   120
         TabIndex        =   73
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label lblUUser 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Add USER"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   6000
         TabIndex        =   72
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H00404000&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   8040
         TabIndex        =   70
         Top             =   2640
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00404000&
         Caption         =   "Role"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   735
         Left            =   8040
         TabIndex        =   67
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label13 
         BackColor       =   &H00404000&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   855
         Left            =   240
         TabIndex        =   65
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label Label12 
         BackColor       =   &H00404000&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   855
         Left            =   120
         TabIndex        =   63
         Top             =   720
         Width           =   2415
      End
      Begin VB.Image Image4 
         Height          =   11580
         Left            =   -4800
         Picture         =   "frmDash.frx":1D2BF
         Top             =   360
         Width           =   25560
      End
   End
   Begin VB.Frame frameNotifications 
      BackColor       =   &H00404000&
      Caption         =   "Notifications"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   8055
      Left            =   6000
      TabIndex        =   44
      Top             =   2760
      Width           =   14295
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   495
         Left            =   10920
         TabIndex        =   60
         Top             =   3360
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   4210688
         CalendarForeColor=   -2147483633
         CalendarTitleBackColor=   8421376
         CalendarTitleForeColor=   -2147483637
         CalendarTrailingForeColor=   -2147483638
         Format          =   158007297
         CurrentDate     =   45835
      End
      Begin VB.ComboBox cmbAAssetID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3840
         TabIndex        =   59
         Top             =   840
         Width           =   3015
      End
      Begin VB.ComboBox cmbATechnicianID 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   10800
         TabIndex        =   57
         Text            =   "SELECT"
         Top             =   2160
         Width           =   3015
      End
      Begin VB.TextBox txtAMessage 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7440
         TabIndex        =   55
         Top             =   5400
         Width           =   6015
      End
      Begin VB.TextBox txtATitle 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10560
         TabIndex        =   53
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtAUserName 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   51
         Top             =   4320
         Width           =   3015
      End
      Begin VB.TextBox txtAUserID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   49
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txtAAssetName 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   46
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404000&
         Caption         =   "MESSAGE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   7680
         TabIndex        =   61
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label lblReply 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADDRESS"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   6120
         TabIndex        =   58
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404000&
         Caption         =   "Technician"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   7560
         TabIndex        =   56
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00404000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   7560
         TabIndex        =   54
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404000&
         Caption         =   "TITLE"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   7440
         TabIndex        =   52
         Top             =   840
         Width           =   2055
      End
      Begin VB.Line Line8 
         BorderWidth     =   5
         X1              =   0
         X2              =   14280
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line6 
         BorderWidth     =   5
         X1              =   7200
         X2              =   7200
         Y1              =   120
         Y2              =   6120
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404000&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   735
         Left            =   360
         TabIndex        =   50
         Top             =   4320
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404000&
         Caption         =   "User ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   360
         TabIndex        =   48
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404000&
         Caption         =   "Asset ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         Caption         =   "Asset Name"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   3015
      End
      Begin VB.Image Image3 
         Height          =   14580
         Left            =   -8400
         Picture         =   "frmDash.frx":3A57E
         Top             =   360
         Width           =   26955
      End
   End
   Begin VB.Frame frameMaintenance 
      BackColor       =   &H00404000&
      Caption         =   "Maintenance"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   8055
      Left            =   6000
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   14295
      Begin VB.CommandButton Command1 
         Caption         =   "REPORT"
         Height          =   615
         Left            =   10920
         TabIndex        =   76
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox txtMDescription 
         BeginProperty Font 
            Name            =   "QuickType II Mono"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6960
         TabIndex        =   41
         Top             =   4560
         Width           =   6975
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   495
         Left            =   10680
         TabIndex        =   40
         Top             =   1800
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   4210688
         CalendarForeColor=   -2147483633
         CalendarTitleBackColor=   8421376
         Format          =   158007297
         CurrentDate     =   45832
      End
      Begin VB.ComboBox cmbMMaintenanceID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10440
         TabIndex        =   38
         Top             =   600
         Width           =   3255
      End
      Begin VB.ComboBox cmbMStatus 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3120
         TabIndex        =   36
         Top             =   4200
         Width           =   2775
      End
      Begin VB.ComboBox cmbMTechnicianID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         ItemData        =   "frmDash.frx":6B9DE
         Left            =   3360
         List            =   "frmDash.frx":6B9E0
         TabIndex        =   34
         Top             =   3000
         Width           =   2895
      End
      Begin VB.ComboBox cmbMAssetID 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3240
         TabIndex        =   33
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox txtMAssetName 
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         TabIndex        =   30
         Top             =   600
         Width           =   3015
      End
      Begin VB.Label lblMAssign 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Assign "
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   5880
         TabIndex        =   43
         Top             =   6840
         Width           =   1815
      End
      Begin VB.Label lblMDescription 
         BackColor       =   &H00404000&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   6960
         TabIndex        =   42
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label lblMDate 
         BackColor       =   &H00404000&
         Caption         =   "Next MAintenance date"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   735
         Left            =   6840
         TabIndex        =   39
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404000&
         Caption         =   "Maintenance ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   7080
         TabIndex        =   37
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label lblMStatus 
         BackColor       =   &H00404000&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         TabIndex        =   35
         Top             =   4200
         Width           =   2055
      End
      Begin VB.Line Line7 
         BorderWidth     =   5
         X1              =   0
         X2              =   14280
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line5 
         BorderWidth     =   5
         X1              =   6720
         X2              =   6720
         Y1              =   240
         Y2              =   6600
      End
      Begin VB.Label lblMTechnician 
         BackColor       =   &H00404000&
         Caption         =   "Technician"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         TabIndex        =   32
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblMAsetID 
         BackColor       =   &H00404000&
         Caption         =   "Asset ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label lblMAssetName 
         BackColor       =   &H00404000&
         Caption         =   "Asset name"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   14580
         Left            =   120
         Picture         =   "frmDash.frx":6B9E2
         Top             =   360
         Width           =   26955
      End
   End
   Begin VB.Frame frameIssueAssets 
      BackColor       =   &H00404000&
      Caption         =   "Assets"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   8055
      Left            =   5880
      TabIndex        =   5
      Top             =   2760
      Width           =   14295
      Begin VB.CommandButton cmdPrint 
         Caption         =   "PRINT"
         Height          =   615
         Left            =   11880
         TabIndex        =   75
         Top             =   6840
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmDash.frx":A694D
         Height          =   7215
         Left            =   -240
         OleObjectBlob   =   "frmDash.frx":A6965
         TabIndex        =   27
         Top             =   240
         Visible         =   0   'False
         Width           =   14535
      End
      Begin VB.ComboBox cmbStatus1 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3360
         TabIndex        =   15
         Text            =   "SELECT"
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox txtAssetName 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3480
         MousePointer    =   3  'I-Beam
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox txtManufacturer 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         MousePointer    =   3  'I-Beam
         TabIndex        =   12
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtSerialNumber 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         MousePointer    =   3  'I-Beam
         TabIndex        =   11
         Top             =   3960
         Width           =   2655
      End
      Begin VB.TextBox txtUserID 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   11760
         MousePointer    =   3  'I-Beam
         TabIndex        =   10
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ComboBox cmbDepartment 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11160
         TabIndex        =   9
         Text            =   "SELECT"
         Top             =   840
         Width           =   2895
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   7920
         TabIndex        =   8
         Top             =   5400
         Width           =   6255
      End
      Begin VB.ComboBox txtAssetID 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   11640
         TabIndex        =   7
         Top             =   4080
         Width           =   2295
      End
      Begin VB.TextBox txtInterval 
         BeginProperty Font 
            Name            =   "QuickType II"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   6
         Text            =   "30"
         Top             =   5160
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   11760
         TabIndex        =   14
         Top             =   1920
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   4210688
         CalendarForeColor=   -2147483633
         CalendarTitleBackColor=   8421376
         Format          =   158007297
         CurrentDate     =   45830
      End
      Begin VB.Label labal 
         BackColor       =   &H00808000&
         Caption         =   "Asset Name"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   26
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00808000&
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label label2 
         BackColor       =   &H00808000&
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label label3 
         BackColor       =   &H00808000&
         Caption         =   "SERIAL Number"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderWidth     =   5
         X1              =   7680
         X2              =   7680
         Y1              =   120
         Y2              =   6360
      End
      Begin VB.Line Line4 
         BorderWidth     =   5
         X1              =   0
         X2              =   14280
         Y1              =   6360
         Y2              =   6360
      End
      Begin VB.Label label3 
         BackColor       =   &H00808000&
         Caption         =   "Deparment"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   0
         Left            =   7920
         TabIndex        =   22
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lbl1 
         BackColor       =   &H00808000&
         Caption         =   "USER ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Index           =   0
         Left            =   7920
         TabIndex        =   21
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label lblIssue 
         Alignment       =   2  'Center
         BackColor       =   &H00404000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ISSUE ASSET"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   6600
         TabIndex        =   20
         Top             =   6720
         Width           =   2295
      End
      Begin VB.Label lbl2 
         BackColor       =   &H00808000&
         Caption         =   "Asset ID"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   615
         Left            =   7920
         TabIndex        =   19
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label lblDescription 
         BackColor       =   &H00808000&
         Caption         =   "Describe Asset Issue"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   12
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   375
         Left            =   7920
         TabIndex        =   18
         Top             =   4920
         Width           =   2535
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00808000&
         Caption         =   "Issue Date"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   495
         Left            =   7920
         TabIndex        =   17
         Top             =   1920
         Width           =   3255
      End
      Begin VB.Label lblInterval 
         BackColor       =   &H00808000&
         Caption         =   "Maintenance Interval (days)"
         BeginProperty Font 
            Name            =   "Perpetua Titling MT"
            Size            =   14.25
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   5040
         Width           =   2775
      End
   End
   Begin VB.Data DatMaintenance 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   18240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Maintenance"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DatNotification 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   18240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Notications"
      Top             =   1320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data DatUsers 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   18360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "User"
      Top             =   1080
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Data DatDepartment 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   18360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Department"
      Top             =   720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data DatAssets 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   18240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Assets"
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data DatLog 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   18240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LoginLogs"
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblAssets 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         ASSETS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   68
      Top             =   2760
      Width           =   5655
   End
   Begin VB.Label lblHead 
      Alignment       =   2  'Center
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AUTOMATED IT ASSET TRACKING AND MAINTENANCE  SYSTEM"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   36
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   2415
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   120
      Width           =   15615
   End
   Begin VB.Line Line2 
      BorderWidth     =   5
      X1              =   5760
      X2              =   5760
      Y1              =   2640
      Y2              =   10920
   End
   Begin VB.Label lblSettings 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         SETTINGS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   9120
      Width           =   5655
   End
   Begin VB.Label lblUSERS 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         USERS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   5655
   End
   Begin VB.Label lblAlerts 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         ALERTS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   5655
   End
   Begin VB.Label lblMaintenance 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "         MAINTENANCE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   855
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   4200
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   2535
      Left            =   0
      Picture         =   "frmDash.frx":A7338
      Top             =   0
      Width           =   2310
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   -240
      X2              =   20040
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Menu mnuAssets 
      Caption         =   "&Assets"
      Visible         =   0   'False
      Begin VB.Menu mnuIssueAsset 
         Caption         =   "Issue Asset"
      End
      Begin VB.Menu mnuViewAssets 
         Caption         =   "View Assets"
      End
   End
End
Attribute VB_Name = "frmDash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Private Sub cmbAAssetID_Click()
    DatUsers.Refresh
    DatNotification.Refresh
    If Not DatAssets.Recordset.BOF Then
        DatAssets.Recordset.MoveFirst
    End If
    If Not DatNotification.Recordset.BOF Then
        DatNotification.Recordset.MoveFirst
    End If
    Do Until DatNotification.Recordset.EOF
            DatAssets.Recordset.MoveFirst
            Do Until DatAssets.Recordset.EOF
                If DatAssets.Recordset.Fields("AssetID") = DatNotification.Recordset.Fields("AssetID") Then
                        DatUsers.RecordSource = "SELECT * FROM User WHERE UserID=" & DatAssets.Recordset.Fields("UserID").Value
                        txtAAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                        txtAUserID.Text = DatAssets.Recordset.Fields("UserID")
                        txtAUserName.Text = DatUsers.Recordset.Fields("Name")
                        txtATitle.Text = DatNotification.Recordset.Fields("Title")
                        DTPicker3.Value = DatNotification.Recordset.Fields("NotificationDate")
                        txtAMessage.Text = DatNotification.Recordset.Fields("Message")
                        DatNotification.Recordset.Edit
                        DatNotification.Recordset.Fields("Status") = "Read"
                        DatNotification.Recordset.Update
                        lblAlerts.BackColor = &H404000

                End If
                DatAssets.Recordset.MoveNext
            Loop
        DatNotification.Recordset.MoveNext
    Loop
        
    
End Sub

Private Sub cmbMAssetID_Click()
    If Not DatAssets.Recordset.BOF Then
        DatAssets.Recordset.MoveFirst
        
    End If
    Do Until DatAssets.Recordset.EOF
        If DatAssets.Recordset.Fields("AssetID") = Val(cmbMAssetID.Text) Then
            txtMAssetName.Text = DatAssets.Recordset.Fields("AssetName")
        End If
        DatAssets.Recordset.MoveNext
    Loop
    If Not DatMaintenance.Recordset.BOF Then
        DatMaintenance.Recordset.MoveFirst
        
    End If
    Do Until DatMaintenance.Recordset.EOF
        If DatMaintenance.Recordset.Fields("AssetID") = Val(cmbMAssetID.Text) Then
            cmbMMaintenanceID.Text = DatMaintenance.Recordset.Fields("MaintenanceID")
            cmbMTechnicianID.Text = DatMaintenance.Recordset.Fields("TechnicianID")
            If Not DatMaintenance.Recordset.Fields("Status") = Null Then
                cmbMStatus.Text = DatMaintenance.Recordset.Fields("Status")
            End If
            If Not DatMaintenance.Recordset.Fields("Description") = Null Then
                   txtDescription.Text = DatMaintenance.Recordset.Fields("Description")
            End If
            DTPicker1.Value = DatMaintenance.Recordset.Fields("NextMaintenance")
        End If
        DatMaintenance.Recordset.MoveNext
        
    Loop
End Sub

Private Sub cmbMMaintenanceID_Click()
   'DatMaintenance.RecordSource = "SELECT * FROM Maintenance WHERE MaintenanceID=" & Val(cmbMMaintenanceID.Text)
   ' DatAssets.RecordSource = "SELECT * FROM Assets WHERE AssetID=" &
    If Not DatAssets.Recordset.BOF Then
        DatAssets.Recordset.MoveFirst
    End If
    If Not DatMaintenance.Recordset.BOF Then
        DatMaintenance.Recordset.MoveFirst
    End If
    Do Until DatMaintenance.Recordset.EOF
        DatAssets.Recordset.MoveFirst
        Do Until DatAssets.Recordset.EOF
            If cmbMMaintenanceID.Text = DatAssets.Recordset.Fields("AssetID") And DatMaintenance.Recordset.Fields("AssetID") = DatAssets.Recordset.Fields("AssetID") Then
                    txtMAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                    cmbMAssetID.Text = DatAssets.Recordset.Fields("AssetID")
                    If Not DatMaintenance.Recordset.Fields("Status") = Null Then
                        cmbMStatus.Text = DatMaintenance.Recordset.Fields("Status")
                    End If
                    If Not DatMaintenance.Recordset.Fields("Description") = Null Then
                        txtMDescription.Text = DatMaintenance.Recordset.Fields("Description")
                    End If
                    DTPicker2.Value = DatMaintenance.Recordset.Fields("NextMaintenance")
                    If DTPicker2.Value < Now Then
                        cmbMStatus.Text = "Missed"
                    End If
                Exit Sub
            End If
            DatAssets.Recordset.MoveNext
        Loop
        DatMaintenance.Recordset.MoveNext
    Loop
    
End Sub



Private Sub cmdLogs_Click()
    If Not Role = "Admin" Then
        Exit Sub
    End If
    Dim UserID As Long
    UserID = Val(InputBox("Enter User's ID"))
    Adodc2.RecordSource = "SELECT * FROM LoginLogs WHERE UserID=" & UserID
    Set DataReport3.DataSource = Adodc2
    DataReport3.Caption = "Login logs"
    DataReport3.Show
    
End Sub

Private Sub Command1_Click()
    Adodc1.RecordSource = "SELECT * FROM Maintenance WHERE AssetID=" & Val(cmbMAssetID.Text)
    Adodc1.Refresh
    Set DataReport2.DataSource = Adodc1
    DataReport2.Caption = "MAINTENANCE REPORT"
    DataReport2.Show
    
End Sub

Private Sub cmdPrint_Click()
    Dim AssetID As Long
    AssetID = Val(InputBox("Enter The Asset Id "))
    
    Adodc1.RecordSource = "SELECT * FROM Assets WHERE AssetID=" & AssetID
    Adodc1.Refresh
    DataReport1.Caption = "ASSET DETAILS"
    Set DataReport1.DataSource = Adodc1
    DataReport1.Show
End Sub

Private Sub Form_Load()
    DatAssets.Refresh
    DatUsers.Refresh
    frameUser.Visible = False
    DatMaintenance.Refresh
    DatNotification.Refresh
    frameIssueAssets.Visible = False
    frameNotifications.Visible = False
    cmbMStatus.AddItem "Maintained"
    cmbMStatus.AddItem "Pending"
    cmbMStatus.AddItem "Missed"
    cmbMStatus.AddItem "Completed"
    cmbMStatus.AddItem "Scheduled"
    DTPicker1.Value = Now
    If Role = "AssetUser" Then
        lblInterval.Visible = False
        txtInterval.Visible = False
        lblMAssign.Visible = False
        DatAssets.RecordSource = "SELECT * FROM Assets WHERE UserID=" & CurrentUserID
        DatUsers.RecordSource = "SELECT * FROM User WHERE UserID=" & CurrentUserID
        lblDate.Caption = "Maintenance Date"
        cmbStatus1.Enabled = False
        DTPicker1.Enabled = False
        cmbDepartment.Enabled = False
        cmbDepartment.Text = ""
        cmbStatus1.Text = ""
        txtUserID.Text = CurrentUserID
        txtUserID.Enabled = False
        
        txtAssetID.Text = "SELECT"
        route = True
        lblIssue.Caption = "SUBMIT"
        If Not DatAssets.Recordset.BOF Then
            DatAssets.Recordset.MoveFirst
        End If
        If Not DatMaintenance.Recordset.BOF Then
            DatMaintenance.Recordset.MoveFirst
        End If

        Do Until DatAssets.Recordset.EOF
            If DatAssets.Recordset.Fields("UserID") = CurrentUserID Then
                txtAssetID.AddItem DatAssets.Recordset.Fields("AssetID")
            End If
            DatAssets.Recordset.MoveNext
        Loop
                
        Do Until DatMaintenance.Recordset.EOF
            DatAssets.Recordset.MoveFirst
            Do Until DatAssets.Recordset.EOF
                If DatMaintenance.Recordset.Fields("AssetID") = DatAssets.Recordset.Fields("AssetID") And DatAssets.Recordset.Fields("UserID") = CurrentUserID Then
                    cmbMMaintenanceID.AddItem DatMaintenance.Recordset.Fields("MaintenanceID").Value
                End If
                DatAssets.Recordset.MoveNext
            Loop
            DatMaintenance.Recordset.MoveNext
            
        Loop
        

        
    End If
    'Admin
    If Role = "Admin" Then
        lblDescription.Visible = False
        txtDescription.Visible = False
        
        cmbStatus1.AddItem "Assigned"
        cmbStatus1.AddItem "Available"
        cmbStatus1.AddItem "Maintenance"
        cmbStatus1.AddItem "Retired"
        cmbDepartment.AddItem "IT Department"
        cmbDepartment.AddItem "Human Resource"
        cmbDepartment.AddItem "Automotive"
        cmbDepartment.AddItem "Finance"
        cmbUDepartment.AddItem "IT Department"
        cmbUDepartment.AddItem "Human Resource"
        cmbUDepartment.AddItem "Automotive"
        cmbUDepartment.AddItem "Finance"
        cmbURole.AddItem "Admin"
        cmbURole.AddItem "AssetUser"
        cmbURole.AddItem "Technician"
        
        
        If Not DatAssets.Recordset.BOF Then
            DatAssets.Recordset.MoveFirst
        End If
        Do Until DatAssets.Recordset.EOF
            txtAssetID.AddItem DatAssets.Recordset.Fields("AssetID")
            cmbMAssetID.AddItem DatAssets.Recordset.Fields("AssetID")

            DatAssets.Recordset.MoveNext
        Loop
        If Not DatUsers.Recordset.BOF Then
            DatUsers.Recordset.MoveFirst
        End If
        

        Do Until DatUsers.Recordset.EOF
            If DatUsers.Recordset.Fields("Role") = "Technician" Then
                cmbMTechnicianID.AddItem DatUsers.Recordset.Fields("UserID")
                cmbATechnicianID.AddItem DatUsers.Recordset.Fields("UserID")
            End If
            DatUsers.Recordset.MoveNext
        Loop
        If Not DatMaintenance.Recordset.BOF Then
            DatMaintenance.Recordset.MoveFirst
        End If
        
        Do Until DatMaintenance.Recordset.EOF
            If DatMaintenance.Recordset.Fields("TechnicianID") = 0 Then
                cmbMMaintenanceID.AddItem DatMaintenance.Recordset.Fields("MaintenanceID").Value
            End If
            DatMaintenance.Recordset.MoveNext
            
        Loop
        If Not DatNotification.Recordset.BOF Then
            DatNotification.Recordset.MoveFirst
        End If
        Do Until DatNotification.Recordset.EOF
            If DatNotification.Recordset.Fields("Status") = "Unread" Then
                cmbAAssetID.AddItem DatNotification.Recordset.Fields("AssetID")
                lblAlerts.BackColor = &H40C0&
                
            End If
            DatNotification.Recordset.MoveNext
        Loop
       
        
    End If
    'Technician
    If Role = "Technician" Then
        lblReply.Visible = False
        DatNotification.Refresh
        lblDescription.Caption = "REQUEST ITEM"
        lblIssue.Caption = "Request"
        cmbStatus1.AddItem "Maintenance"
        lblMAssign.Caption = "SUBMIT"
        DatMaintenance.Refresh
        If Not DatAssets.Recordset.BOF Then
            DatAssets.Recordset.MoveFirst
        End If
        If Not DatMaintenance.Recordset.BOF Then
            DatMaintenance.Recordset.MoveFirst
            
        End If
        Do Until DatMaintenance.Recordset.EOF
            If Not DatAssets.Recordset.BOF Then
                DatAssets.Recordset.MoveFirst
            End If

            Do Until DatAssets.Recordset.EOF
                If DatMaintenance.Recordset.Fields("TechnicianID") = CurrentUserID Then
                    If DatMaintenance.Recordset.Fields("AssetID") = DatAssets.Recordset.Fields("AssetID") Then
                        txtAssetID.AddItem DatAssets.Recordset.Fields("AssetID")
                        cmbMMaintenanceID.AddItem DatMaintenance.Recordset.Fields("MaintenanceID").Value
                    
                    End If
                End If
                DatAssets.Recordset.MoveNext
                
            
            Loop
           DatMaintenance.Recordset.MoveNext
        Loop
         If Not DatNotification.Recordset.BOF Then
            DatNotification.Recordset.MoveFirst
        End If
        Do Until DatNotification.Recordset.EOF
            If DatNotification.Recordset.Fields("UserID") = CurrentUserID Then
                cmbAAssetID.AddItem DatNotification.Recordset.Fields("AssetID")
                
            End If
            DatNotification.Recordset.MoveNext
        Loop
        
    End If
    
End Sub







Private Sub lblAlerts_Click()
    frameIssueAssets.Visible = False
    frameMaintenance.Visible = False
    frameUser.Visible = False
    If frameNotifications.Visible = False Then
        frameNotifications.Visible = True
    Else
        frameNotifications.Visible = False
    End If
End Sub

Private Sub lblAssets_Click(Index As Integer)
    frameMaintenance.Visible = False
    frameNotifications.Visible = False
    frameUser.Visible = False
    If Role = "Admin" Then
        PopupMenu mnuAssets
        Exit Sub
    End If
    If frameIssueAssets.Visible = True Then
        frameIssueAssets.Visible = False
        
    Else
        frameIssueAssets.Visible = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    DatLog.RecordSource = "SELECT * FROM LoginLogs WHERE LogID=" & LogID
    DatLog.Refresh
    DatUsers.Refresh
    DatUsers.RecordSource = "SELECT * FROM User WHERE UserID=" & CurrentUserID
    If Not DatLog.Recordset.EOF Then
        DatLog.Recordset.Edit
        DatLog.Recordset.Fields("LogoutTime") = Now
        DatLog.Recordset.Update
        
    End If
    If Not DatUsers.Recordset.EOF Then
            DatUsers.Recordset.Edit
            DatUsers.Recordset.Fields("Status") = "Offline"
            DatUsers.Recordset.Update
     End If
    
End Sub



Private Sub lblIssue_Click()
    Dim AssetName As String
    Dim Manufacturer As String
    Dim status As String
    Dim serialNo As String
    Dim Department As String
    Dim purchaseDate As Date
    Dim UserID As Long
    Dim max_id As Long
    Dim DeptID As Long
    Dim num As Integer
    Dim maxid As Long
    
    max_id = 0
    If Role = "Admin" Then
        DatMaintenance.Refresh
        DatUsers.Refresh
        DatAssets.Refresh
        DatDepartment.Refresh
        num = Val(txtInterval.Text)
        AssetName = frmDash.txtAssetName.Text
        Manufacturer = txtManufacturer.Text
        status = cmbStatus1.Text
        maxid = 0
        serialNo = txtSerialNumber.Text
        Department = cmbDepartment.Text
        purchaseDate = DTPicker1.Value
        UserID = Val(txtUserID.Text)
        If AssetName = "" Or num = Null Or num = 0 Or Manufacturer = "" Or status = "SELECT" Or serialNo = "" Or Department = "SELECT" Or status = "" Or Department = "" Or UserID = Null Or UserID = 0 Then
            MsgBox "All Fields Are Required", vbExclamation
            Exit Sub
        End If

        DatDepartment.RecordSource = "SELECT * FROM Department WHERE DepartmentName='" & Department & "'"
        If Not DatMaintenance.Recordset.BOF Then
            DatMaintenance.Recordset.MoveFirst
            
        End If
        Do Until DatMaintenance.Recordset.EOF
            If DatMaintenance.Recordset.Fields("MaintenanceID") > maxid Then
                maxid = DatMaintenance.Recordset.Fields("MaintenanceID")
            End If
            DatMaintenance.Recordset.MoveNext
            
            
        Loop
        If purchaseDate > Now Then
            MsgBox "Enter a valid date or set Date and time", vbCritical
            Exit Sub
        End If
        If route = False Then
            DatMaintenance.RecordSource = "SELECT * FROM Maintenance WHERE AssetID=" & DatAssets.Recordset.Fields("AssetID")
            DatAssets.Recordset.Edit
            DatAssets.Recordset.Fields("AssetName") = AssetName
            DatAssets.Recordset.Fields("Manufacturer") = Manufacturer
            DatAssets.Recordset.Fields("Status") = status
            DatAssets.Recordset.Fields("SerialNumber") = serialNo
            DatAssets.Recordset.Fields("DepartmentID") = DatDepartment.Recordset.Fields("DepartmentID")
            DatAssets.Recordset.Fields("PurchaseDate") = purchaseDate
            DatAssets.Recordset.Fields("UserID") = UserID
            DatAssets.Recordset.Update
            DatMaintenance.Recordset.Edit
            DatMaintenance.Recordset.Fields("MaintenanceInterval") = num
            DatMaintenance.Recordset.Fields("NextMaintenance") = DateAdd("d", num, purchaseDate)
            DatMaintenance.Recordset.Update
            MsgBox "Asset  Of id:" & DatAssets.Recordset.Fields("AssetID") & " Updated successfully", vbInformation
            txtAssetName.Text = ""
            txtManufacturer.Text = ""
            txtSerialNumber.Text = ""
            txtUserID.Text = ""
            txtAssetID.Text = ""
            Exit Sub
        End If
        DatMaintenance.Recordset.AddNew
        DatMaintenance.Recordset.Fields("MaintenanceID") = maxid + 1
        DatMaintenance.Recordset.Fields("AssetID") = i
        DatMaintenance.Recordset.Fields("MaintenanceInterval") = num
        DatMaintenance.Recordset.Fields("NextMaintenance") = DateAdd("d", num, purchaseDate)
        DatMaintenance.Recordset.Fields("TechnicianID") = 0
        DatMaintenance.Recordset.Update
        DatAssets.Recordset.AddNew
        DatAssets.Recordset.Fields("AssetName") = AssetName
        DatAssets.Recordset.Fields("Manufacturer") = Manufacturer
        DatAssets.Recordset.Fields("Status") = status
        DatAssets.Recordset.Fields("SerialNumber") = serialNo
        DatAssets.Recordset.Fields("DepartmentID") = DatDepartment.Recordset.Fields("DepartmentID")
        DatAssets.Recordset.Fields("PurchaseDate") = purchaseDate
        DatAssets.Recordset.Fields("UserID") = UserID
        DatAssets.Recordset.Fields("AssetID") = i
        DatAssets.Recordset.Update
    
    
        MsgBox "Asset Issuance Success!", vbInformation
        Unload Me
        Load frmDash
        frameIssueAssets.Visible = True
        
        frmDash.Caption = Role
        frmDash.Show
        
        txtAssetName.Text = ""
        txtManufacturer.Text = ""
        txtSerialNumber.Text = ""
        txtUserID.Text = ""
        txtAssetID.Text = ""
 
    End If
    If Role = "AssetUser" Then
        DatNotification.Refresh
        If txtAssetName.Text = "" Or txtManufacturer = "" Then
            MsgBox "Choose Asset To Report", vbExclamation
            Exit Sub
        End If
        If txtDescription.Text = "" Then
            MsgBox "Enter Description!", vbExclamation
            Exit Sub
        End If
        If Not DatNotification.Recordset.BOF Then
            DatNotification.Recordset.MoveFirst
        End If
        Do Until DatNotification.Recordset.EOF
            If DatNotification.Recordset.Fields("NotificationID") > max_id Then
                max_id = DatNotification.Recordset.Fields("NotificationID")
            End If
            DatNotification.Recordset.MoveNext
        Loop
        DatNotification.Recordset.AddNew
        DatNotification.Recordset.Fields("Title") = txtAssetName.Text & "    Issue"
        DatNotification.Recordset.Fields("Message") = txtDescription.Text
        DatNotification.Recordset.Fields("Status") = "Unread"
        DatNotification.Recordset.Fields("NotificationDate") = Now
        DatNotification.Recordset.Fields("NotificationID") = max_id + 1
        'DatNotification.Recordset.Fields("UserID") = CurrentUserID
        DatNotification.Recordset.Fields("AssetID") = txtAssetID.Text
        DatNotification.Recordset.Update
        MsgBox "Issue Reported Successfully!", vbInformation
        Unload Me
        Load frmDash
        frmDash.Caption = Role
        frmDash.Show
        
        txtAssetName.Text = ""
        txtManufacturer.Text = ""
        txtSerialNumber.Text = ""
        txtUserID.Text = ""
        txtAssetID.Text = ""

    End If
    If Role = "Technician" Then
        If txtAssetName.Text = "" Or txtManufacturer = "" Or txtAssetID.Text = 0 Or txtAssetID.Text = "" Or txtDescription.Text = "" Then
            MsgBox "Choose Asset To Report", vbExclamation
            Exit Sub
        End If

        DatNotification.Recordset.AddNew
        DatNotification.Recordset.Fields("Title") = txtAssetName.Text & "Issue"
        DatNotification.Recordset.Fields("Message") = txtDescription.Text
        DatNotification.Recordset.Fields("Status") = "Unread"
        DatNotification.Recordset.Fields("NotificationDate") = Now
        DatNotification.Recordset.Fields("NotificationID") = max_id + 1
        'DatNotification.Recordset.Fields("UserID") = CurrentUserID
        DatNotification.Recordset.Fields("AssetID") = txtAssetID.Text
        DatNotification.Recordset.Update
        MsgBox "Requested Successfully!", vbInformation
        Unload Me
        Load frmDash
        frmDash.Caption = Role
        frmDash.Show
        
        txtAssetName.Text = ""
        txtManufacturer.Text = ""
        txtSerialNumber.Text = ""
        txtUserID.Text = ""
        txtAssetID.Text = ""
 
    End If
End Sub




Private Sub lblMaintenance_Click(Index As Integer)
    frameIssueAssets.Visible = False
    frameNotifications.Visible = False
    frameUser.Visible = False
    If frameMaintenance.Visible = False Then
        frameMaintenance.Visible = True
    Else
        frameMaintenance.Visible = False
    End If
End Sub

Private Sub lblMAssign_Click()
    If Role = "Admin" Then
        If cmbMTechnicianID.Text = "" Or cmbMMaintenanceID.Text = "" Then
            MsgBox "Enter Required Fields To Assign", vbExclamation
            
            Exit Sub
        End If
        DatMaintenance.Recordset.Edit
        If Not txtMDescription.Text = "" Then
            DatMaintenance.Recordset.Fields("Description") = txtMDescription
        End If
        If Not cmbMStatus.Text = "" Then
            DatMaintenance.Recordset.Fields("Status") = cmbMStatus.Text
        End If
        DatMaintenance.Recordset.Fields("TechnicianID") = cmbMTechnicianID
        DatMaintenance.Recordset.Update
        MsgBox "Task Assigned", vbInformation
        Unload Me
        Load frmDash
        frmDash.Caption = Role
        frameMaintenance.Visible = True
        frmDash.Caption = Role
        frmDash.Show
        
             
    End If
    If Role = "Technician" Then
        If cmbMAssetID.Text = "" Or txtMDescription.Text = "" Or cmbMMaintenanceID.Text = "" Or cmbMStatus.Text = "" Then
                MsgBox "Enter data to Proceed", vbExclamation
                Exit Sub
        End If
        DatMaintenance.Recordset.Edit
        DatMaintenance.Recordset.Fields("Description") = txtMDescription
        DatMaintenance.Recordset.Fields("Status") = cmbMStatus.Text
        DatMaintenance.Recordset.Fields("TechnicianID") = CurrentUserID
        DatMaintenance.Recordset.Update
        MsgBox "Task Assigned", vbInformation
        Unload Me
        Load frmDash
        frmDash.Caption = Role
        frameMaintenance.Visible = True
        
        frmDash.Show

    End If
End Sub

Private Sub lblReply_Click()
    'Assign to Technician
    If cmbAAssetID.Text = "" Or cmbATechnicianID.Text = "" Then
        MsgBox "Select an asset  to Address Or technician to foward to", vbExclamation
        Exit Sub
    End If
    DatNotification.Refresh
    DatNotification.RecordSource = "SELECT* FROM Notifications WHERE AssetID=" & Val(cmbAAssetID.Text)
    DatNotification.Recordset.Edit
    DatNotification.Recordset.Fields("UserID") = Val(cmbATechnicianID.Text)
    DatNotification.Recordset.Fields("Message") = txtAMessage.Text
    DatNotification.Recordset.Update
    MsgBox "Message Fowarded", vbInformation
    
    
End Sub

Private Sub lblSettings_Click()
    frameMaintenance.Visible = False
    If frameUser.Caption = "USER" Then
        frameUser.Caption = "SETTINGS"
    End If
    
    frameIssueAssets.Visible = False
    frameNotifications.Visible = False
    If frameUser.Visible = True And frameUser.Caption = "USER" Then
        frameUser.Visible = False
    Else
        frameUser.Visible = True
        DatUsers.Refresh
        If Not DatUsers.Recordset.BOF Then
            DatUsers.Recordset.MoveFirst
        End If
        Do Until DatUsers.Recordset.EOF
            If DatUsers.Recordset.Fields("UserID") = CurrentUserID Then
                txtUUserName.Text = DatUsers.Recordset.Fields("Name")
                txtUUserID.Text = DatUsers.Recordset.Fields("UserID")
                'cmbUDepartment.Text = DatUsers.Recordset.Fields("Department")
                cmbURole.Text = DatUsers.Recordset.Fields("Role")
                txtUPassword.Text = DatUsers.Recordset.Fields("Password")
                lblUUser.Caption = "EDIT DETAILS"
                Exit Sub
            End If
            DatUsers.Recordset.MoveNext
            
                
        Loop
    End If
    
    
End Sub

Private Sub lblUSERS_Click()
    frameIssueAssets.Visible = False
    frameMaintenance.Visible = False
    frameNotifications.Visible = False
    frameUser.Caption = "USER"
    lblUUser.Caption = "ADD USER"
    If frameUser.Visible = True And frameUser.Caption = "SETTINGS" Then
        frameUser.Visible = False
    Else
        frameUser.Visible = True
        txtUUserName.Text = ""
        txtUPassword.Text = ""

    End If
    
End Sub

Private Sub lblUUser_Click()
    If txtUPassword.Text = "" Then
        MsgBox "Enter Password!", vbExclamation
        Exit Sub
    End If
    Dim departmentid As Long
    If Role = "Admin" And frameUser.Caption = "USER" Then
        If txtUUserName = "" Or txtUPassword.Text = "" Or cmbURole.Text = "" Or cmbURole.Text = "SELECT" Then
            MsgBox "All Fields Are Required", vbExclamation
            Exit Sub
        End If
        If Not DatDepartment.Recordset.BOF Then
            DatDepartment.Recordset.MoveFirst
        End If
        Do Until DatDepartment.Recordset.EOF
            If DatDepartment.Recordset.Fields("DepartmentName") = cmbUDepartment.Text Then
                departmentid = DatDepartment.Recordset.Fields("DepartmentID")
            End If
            DatDepartment.Recordset.MoveNext
        Loop
        DatUsers.Recordset.AddNew
        DatUsers.Recordset.Fields("Name") = txtUUserName.Text
        DatUsers.Recordset.Fields("Password") = txtUPassword.Text
        DatUsers.Recordset.Fields("Status") = Offline
        DatUsers.Recordset.Fields("Role") = txtURole.Text
        DatUsers.Recordset.Fields("UserID") = txtUUserID.Text
        DatUsers.Recordset.Fields("DepartmentID") = departmentid
        DatUsers.Recordset.Update
        MsgBox "User Added Successfully", vbInformation
        txtUUserName.Text = ""
        txtUPassword.Text = ""
        Exit Sub
    End If
    DatUsers.Recordset.Edit
    DatUsers.Recordset.Fields("Password") = txtUPassword.Text
    DatUsers.Recordset.Update
    MsgBox "Password Changed!", vbInformation
    
        
    
End Sub

Private Sub txtAssetID_Change()
    If Role = "Admin" Then
        If IsNumeric(txtAssetID.Text) And route = False Then
            Dim AssetID As Long
            AssetID = txtAssetID.Text
            If Not DatAssets.Recordset.BOF Then
                DatAssets.Recordset.MoveFirst
            End If
            If Not AssetID = 0 Then
                Do Until DatAssets.Recordset.EOF
                    If DatAssets.Recordset.Fields("AssetID") = AssetID Then
                        DatDepartment.RecordSource = "SELECT * FROM Department WHERE DepartmentID=" & DatAssets.Recordset.Fields("DepartmentID")
                        txtManufacturer.Text = DatAssets.Recordset.Fields("Manufacturer")
                        cmbDepartment.Text = DatDepartment.Recordset.Fields("DepartmentName")
                        cmbStatus1.Text = DatAssets.Recordset.Fields("Status")
                        txtAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                        txtSerialNumber.Text = DatAssets.Recordset.Fields("SerialNumber")
                        DTPicker1.Value = DatAssets.Recordset.Fields("PurchaseDate")
                        txtUserID.Text = DatAssets.Recordset.Fields("UserID")
                        Exit Sub
                    End If
                    DatAssets.Recordset.MoveNext
                Loop
                route = False
        
                MsgBox "No asset With ID: " & AssetID, vbExclamation
                txtAssetName.Text = ""
                txtManufacturer.Text = ""
                txtSerialNumber.Text = ""
                txtUserID.Text = ""
                txtAssetID.Text = ""

            End If
        End If
    End If
    If Role = "AssetUser" And route = True Then
        MsgBox "Select an Asset On the drop Down", vbInformation
    End If
End Sub

Private Sub txtAssetName_Change()
    If Role = "Admin" Then

        If txtAssetID.Text = "" Then
            route = True
    
            DatAssets.Refresh
            i = 0
            If Not DatAssets.Recordset.BOF Then
                DatAssets.Recordset.MoveFirst
            End If
            Do Until DatAssets.Recordset.EOF
                If DatAssets.Recordset.Fields("AssetID") > i Then
                    i = DatAssets.Recordset.Fields("AssetID")
                End If
                DatAssets.Recordset.MoveNext
        
            Loop
            i = i + 1
            txtAssetID.Text = i
        End If
        If txtAssetName.Text = "" And route = True Then
            txtAssetID.Text = ""
            route = False
        End If
    End If
    'If Role = "AssetUser" Or Role = "Technician" Then
    '    MsgBox "Select an Asset On the drop Down", vbInformation
   ' End If


End Sub
Private Sub txtAssetID_Click()
If Role = "Admin" Then
        If IsNumeric(txtAssetID.Text) And route = False Then
            Dim AssetID As Long
            AssetID = Val(txtAssetID.Text)
            If Not DatAssets.Recordset.BOF Then
                DatAssets.Recordset.MoveFirst
            End If
            If Not AssetID = 0 Then
                Do Until DatAssets.Recordset.EOF
                    If DatAssets.Recordset.Fields("AssetID") = AssetID Then
                        DatDepartment.RecordSource = "SELECT * FROM Department WHERE DepartmentID=" & DatAssets.Recordset.Fields("DepartmentID")
                        txtManufacturer.Text = DatAssets.Recordset.Fields("Manufacturer")
                        cmbDepartment.Text = DatDepartment.Recordset.Fields("DepartmentName")
                        cmbStatus1.Text = DatAssets.Recordset.Fields("Status")
                        txtAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                        txtSerialNumber.Text = DatAssets.Recordset.Fields("SerialNumber")
                        DTPicker1.Value = DatAssets.Recordset.Fields("PurchaseDate")
                        txtUserID.Text = DatAssets.Recordset.Fields("UserID")
                        Exit Sub
                    End If
                    DatAssets.Recordset.MoveNext
                Loop
                route = False
        
                MsgBox "No asset With ID: " & AssetID, vbExclamation
                txtAssetName.Text = ""
                txtManufacturer.Text = ""
                txtSerialNumber.Text = ""
                txtUserID.Text = ""
                txtAssetID.Text = ""

            End If
        End If
    End If
    If Role = "AssetUser" Then
        If Not DatAssets.Recordset.BOF Then
            DatAssets.Recordset.MoveFirst
        End If
        Do Until DatAssets.Recordset.EOF
            If DatAssets.Recordset.Fields("AssetID") = Val(txtAssetID.Text) Then
                txtManufacturer.Text = DatAssets.Recordset.Fields("Manufacturer")
                cmbDepartment.Text = DatDepartment.Recordset.Fields("DepartmentName")
                cmbStatus1.Text = DatAssets.Recordset.Fields("Status")
                txtAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                txtSerialNumber.Text = DatAssets.Recordset.Fields("SerialNumber")
                DTPicker1.Value = DatAssets.Recordset.Fields("PurchaseDate")
                txtUserID.Text = DatAssets.Recordset.Fields("UserID")
            End If
            DatAssets.Recordset.MoveNext
        Loop
        
    End If
    If Role = "Technician" Then
        DatAssets.Refresh
        If Not DatAssets.Recordset.BOF Then
            DatAssets.Recordset.MoveFirst
        End If
        Do Until DatAssets.Recordset.EOF
            If DatAssets.Recordset.Fields("AssetID") = Val(txtAssetID.Text) Then
                txtManufacturer.Text = DatAssets.Recordset.Fields("Manufacturer")
                cmbDepartment.Text = DatDepartment.Recordset.Fields("DepartmentName")
                cmbStatus1.Text = DatAssets.Recordset.Fields("Status")
                txtAssetName.Text = DatAssets.Recordset.Fields("AssetName")
                txtSerialNumber.Text = DatAssets.Recordset.Fields("SerialNumber")
                DTPicker1.Value = DatAssets.Recordset.Fields("PurchaseDate")
                txtUserID.Text = DatAssets.Recordset.Fields("UserID")
            End If
            DatAssets.Recordset.MoveNext
        Loop
        
    End If


End Sub

Private Sub txtInterval_Change()
    If Not IsNumeric(txtInterval.Text) Then
            MsgBox "Enter An Integer", vbExclamation
            txtInterval.Text = ""
            Exit Sub
    End If
End Sub

Private Sub txtUserID_Change()
    
    If Role = "AssetUser" Then
        Exit Sub
    End If
    If IsNumeric(txtUserID.Text) Then
        If Not DatUsers.Recordset.BOF Then
            DatUsers.Recordset.MoveFirst
        End If
        If Not DatDepartment.Recordset.BOF Then
            DatDepartment.Recordset.MoveFirst
            
        End If
        Do Until DatUsers.Recordset.EOF
            If DatUsers.Recordset.Fields("UserID") = Val(txtUserID.Text) Then
                Do Until DatDepartment.Recordset.EOF
                    If DatDepartment.Recordset.Fields("DepartmentID") = DatUsers.Recordset.Fields("DepartmentID") Then
                                                cmbDepartment.Text = DatDepartment.Recordset.Fields("DepartmentName")
                        Exit Sub
                    End If
                    DatDepartment.Recordset.MoveNext
                Loop
            End If
            DatUsers.Recordset.MoveNext
        Loop
        MsgBox "User with Id:" & txtUserID.Text & " Not Found", vbExclamation
        
        
    End If
End Sub
Private Sub mnuIssueAsset_Click()
        DBGrid1.Visible = False
        frameIssueAssets.Visible = True
End Sub
Private Sub mnuViewAssets_Click()
    DBGrid1.Visible = True
    
    frameIssueAssets.Visible = True
    
    

End Sub

Private Sub txtUUserName_Change()
    If frameUser.Caption = "USER" Then
        If txtUUserName.Text = "" Then
            txtUUserID.Text = ""
            Exit Sub
        End If
    
        Dim maxid As Long
        maxid = 0
        DatUsers.Refresh
        If Not DatUsers.Recordset.BOF Then
            DatUsers.Recordset.MoveFirst
        End If
        Do Until DatUsers.Recordset.EOF
            If DatUsers.Recordset.Fields("UserID") > maxid Then
                maxid = DatUsers.Recordset.Fields("UserID")
            End If
            DatUsers.Recordset.MoveNext
        Loop
        txtUUserID.Text = maxid
        Exit Sub
    End If
    
    
End Sub
