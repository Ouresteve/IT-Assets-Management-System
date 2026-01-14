VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOGIN"
   ClientHeight    =   5115
   ClientLeft      =   6540
   ClientTop       =   2265
   ClientWidth     =   7950
   BeginProperty Font 
      Name            =   "Perpetua Titling MT"
      Size            =   14.25
      Charset         =   0
      Weight          =   300
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   5115
   ScaleWidth      =   7950
   Begin VB.Data DatLog 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "LoginLogs"
      Top             =   4200
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data DatUser 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\Duncons\Desktop\Steve\ITAsset.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   435
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "User"
      Top             =   4560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "QuickType II"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   4560
      MousePointer    =   3  'I-Beam
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   3015
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "QuickType"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4440
      MousePointer    =   3  'I-Beam
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label lblLogin 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "    login"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PASSWORD"
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "USER NAME"
      ForeColor       =   &H8000000F&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Image Image3 
      Height          =   2610
      Left            =   4200
      Picture         =   "frmLogin.frx":2506
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Image Image2 
      Height          =   2610
      Left            =   0
      Picture         =   "frmLogin.frx":4A0C
      Top             =   2520
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   2610
      Left            =   4200
      Picture         =   "frmLogin.frx":6F12
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Sub lblLogin_Click()
    Dim Name As String
    Dim Pass As String
    Dim max_num As Integer
    Dim CurrentTime As Date
    Dim username As String
    Dim computername As String

    username = Environ("USERNAME")
    computername = Environ$("COMPUTERNAME")
    
    Name = txtName.Text
    Pass = txtPass.Text
    DatUser.Refresh
    DatLog.Refresh
    
    max_num = 0
    If Name = "" Or Pass = "" Then
        MsgBox "Enter All Login Credentials to Login!", vbExclamation
        
        Exit Sub
    End If
    If Not DatLog.Recordset.BOF Then
    
        DatLog.Recordset.MoveFirst
    End If
    If Not DatUser.Recordset.BOF Then
                DatUser.Recordset.MoveFirst
    End If
    Do Until DatLog.Recordset.EOF
        If DatLog.Recordset.Fields("LogID") > max_num Then
            max_num = DatLog.Recordset.Fields("LogID")
        End If

        DatLog.Recordset.MoveNext
    Loop
    Do Until DatUser.Recordset.EOF
        If DatUser.Recordset.Fields("Name") = Name And DatUser.Recordset.Fields("Password") = Pass Then
            
            LogID = max_num + 1
            Role = DatUser.Recordset.Fields("Role")
            CurrentUserID = DatUser.Recordset.Fields("UserID")
            frmDash.Caption = Role & ": " & Name
            DatLog.Recordset.AddNew
            DatLog.Recordset.Fields("LogID") = LogID
            DatLog.Recordset.Fields("LoginTime") = Now
            DatLog.Recordset.Fields("ComputerName") = computername
            DatLog.Recordset.Fields("UserID") = CurrentUserID
            DatLog.Recordset.Update
            Load frmDash
            frmDash.Show
            DatUser.Recordset.Edit
            DatUser.Recordset.Fields("Status") = "Online"
            DatUser.Recordset.Update
            
            MsgBox "Login Success!", vbInformation
            Unload Me
            Exit Sub
        End If
        DatUser.Recordset.MoveNext
    Loop
    MsgBox "Log in Failed: Wrong Password Or User Name!", vbExclamation
    
End Sub
