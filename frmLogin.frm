VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inviare il CodiceInstallazione a support@alfierisoftware.net"
   ClientHeight    =   3240
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1914.298
   ScaleMode       =   0  'User
   ScaleWidth      =   6168.875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmail 
      BackColor       =   &H0080FFFF&
      Caption         =   "Send Code"
      Height          =   735
      Left            =   5280
      Picture         =   "frmLogin.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Invia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      Picture         =   "frmLogin.frx":0719
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtlicenza 
      Height          =   375
      Left            =   1920
      TabIndex        =   15
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H0080FF80&
      Caption         =   "Esci"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Picture         =   "frmLogin.frx":2B0B
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FF8080&
      Caption         =   "Salva"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frmLogin.frx":484D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtcodice 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   480
      Width           =   3015
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2880
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   3000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2880
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtlicence 
      DataField       =   "Pagamento"
      DataSource      =   "Adodc5"
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Text            =   "licence"
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Prova"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLogin.frx":6C3F
      Height          =   255
      Left            =   4320
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   450
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      ColumnCount     =   17
      BeginProperty Column00 
         DataField       =   "Order Code"
         Caption         =   "Order Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Date"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cliente"
         Caption         =   "cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Address"
         Caption         =   "Address"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Tire Code"
         Caption         =   "Tire Code"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Tire Description"
         Caption         =   "Tire Description"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Quantity"
         Caption         =   "Quantity"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "numerorighe"
         Caption         =   "numerorighe"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Numero DDT"
         Caption         =   "Numero DDT"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Numero Fattura"
         Caption         =   "Numero Fattura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Numero Preventivo"
         Caption         =   "Numero Preventivo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "campo1"
         Caption         =   "campo1"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "campo2"
         Caption         =   "campo2"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "campo3"
         Caption         =   "campo3"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "ImportoTotale"
         Caption         =   "ImportoTotale"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column15 
         DataField       =   "TotaleIva"
         Caption         =   "TotaleIva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "TotaleImponibile"
         Caption         =   "TotaleImponibile"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   887,371
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1098,7
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1633,677
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1633,677
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   3045
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   4320
      Top             =   2040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "ricdummy"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   4320
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from cassa"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   495
      Left            =   4320
      Top             =   960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from cassa"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   615
      Left            =   1800
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Azienda"
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   5040
      Top             =   960
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from customers"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Codice Installazione"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Licenza N."
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1800
   End
   Begin VB.Label Label3 
      Caption         =   "MODULO MULTILINGUA"
      Height          =   255
      Left            =   360
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Per ricevere il codice di attivazione inviare il codice d installazione all'email support@alfierisoftware.net "
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   2160
      Width           =   5055
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   15
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codice di Attivazione"
      Height          =   390
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1800
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()

End
    
End Sub


Private Sub cmdOK_Click()
    'Verifica la validità della password
    Dim a, b, c As Integer
    Dim stringa, stringa1, stringa2 As String
    
    stringa = "select * from cassa"
    stringa1 = "select * from customers"
    stringa2 = "select * from anamnesi"
    With Adodc2
     .RecordSource = stringa
     .Refresh
     End With
     
     With Adodc3
     .RecordSource = stringa1
     .Refresh
     End With
    
     With Adodc4
     .RecordSource = stringa2
     .Refresh
     End With
    
    
    Select Case Right(txtPassword, Len(Text1.Text))
    
      Case Is = Text1.Text
       
        'Inserire qui il codice per passare alla subroutine chiamante un valore che indica
        'che la password è valida. Il modo più semplice è l'impostazione di una variabile globale.
        LoginSucceeded = True
        Text1.Text = ""
        
        
     
        txtlicence = winproductid
        Adodc5.Recordset!pagamento = txtPassword
       Adodc5.Recordset.Update
       
     
        
        
        DataGrid1.Columns(1) = ""
        Adodc1.Recordset.Update
        
        Me.Hide
    
     Case Is = "XX"
     
     a = Adodc2.Recordset.RecordCount
     b = Adodc3.Recordset.RecordCount
     c = Adodc4.Recordset.RecordCount
     If a > 50 Or b > 12 Or c > 50 Then
     
 MsgBox messaggi(0), , messaggi(1)
 
      '  txtPassword.SetFocus
       
        Me.Hide
        
        End

End If

 MsgBox messaggi(2), , messaggi(1)
 
       'txtPassword.SetFocus
       

       Me.Hide
    
    
       Case Else
        MsgBox messaggi(3), , messaggi(1)
        
        'txtPassword.SetFocus
        
        Me.Hide
        End
    End Select
    
End Sub

Private Sub Command1_Click()
txtPassword = "XX"
cmdOK_Click
End Sub

Private Sub Command2_Click()


If Len(txtlicenza) <> 12 Then GoTo dopo

 varia1 = (Left(txtlicenza, 4))
varia2 = (Mid(txtlicenza, 5, 4))
varia3 = (Right(txtlicenza, 4))

 
 'If varia2 <> varia5 Then GoTo dopo ' controllo programma
 
If Abs(varia1 - varia2) = 8999 - varia3 Then



lblLabels(0).Visible = True
lblLabels(1).Visible = True
txtPassword.Visible = True

txtcodice = Mid(varia1, 3, 2) + txtcodice + Mid(varia2, 3, 2)

txtcodice.Visible = True
cmdOK.Visible = True
'cmdmail.Visible = True

command2.Visible = False

 Else

 End If
 
 
 Exit Sub
 
dopo:
 MsgBox messaggi(3), , messaggi(1)
        
End Sub

Private Sub Command3_Click()


End Sub

Private Sub Form_Load()

Dim a As String

lblLabels(0).Visible = False
lblLabels(1).Visible = False
txtPassword.Visible = False
txtcodice.Visible = False
cmdOK.Visible = False
cmdmail.Visible = False

 Dim SerialNumber As Long
    
    'Get The Computer Name in the registry
    'StartSysInfo
    SerialNumber = GetSerialNumber("c:\")
   ' SysInfoPath = Str(SerialNumber)

 
lingue


 varia5 = 1166 ' numero di codice del programma
  
        winproductid = GetStringValue("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ProductId")
        txtlicence = winproductid
        
        
    Text2 = Mid(txtlicence, 2, 1) + Mid(txtlicence, 5, 1) + Mid(txtlicence, 12, 1) + Mid(txtlicence, 14, 1) + Mid(txtlicence, 16, 1)
    Text3 = Right(txtlicence, 1)
    

      
      
      
      Select Case Val(Text2)
          Case Is < 10000
         txtcodice = "1qscvhu8"
          Text1.Text = "2wdvbjo0"
       Case Is < 14000
         txtcodice = "2wdvbjo0"
         Text1.Text = "3efbnkop"
        Case Is < 18000
         txtcodice = "3efbnkop"
         Text1.Text = "9uhvcde3"
       Case Is < 22000
         txtcodice = "9uhvcde3"
         Text1.Text = "0polk87yh"
       Case Is < 26000
         txtcodice = "0polk87yh"
         Text1.Text = "njhgferrq"
        Case Is < 30000
         txtcodice = "njhgferrq"
         Text1.Text = "hfgjur874"
          Case Is < 34000
         txtcodice = "hfgjur874"
         Text1.Text = "lkdrfcvnm"
       Case Is < 38000
         txtcodice = "lkdrfcvnm"
         Text1.Text = "rtyyujdkwj"
        Case Is < 42000
         txtcodice = "rtyyujdkwj"
         Text1.Text = "eireiofhf"
       Case Is < 46000
         txtcodice = "eireiofhf"
         Text1.Text = "hfaoehhfii"
       Case Is < 50000
         txtcodice = "hfaoehhfii"
         Text1.Text = "fjojrrijgjr"
        Case Is < 54000
         txtcodice = "fjojrrijgjr"
         Text1.Text = "1qscvhu8"





        Case Is < 58000
         txtcodice = "N1qscvhu8"
          Text1.Text = "2wdvAbjo0"
       Case Is < 62000
         txtcodice = "N2wdvbjo0"
         Text1.Text = "3efbAnkop"
        Case Is < 66000
         txtcodice = "N3efbnkop"
         Text1.Text = "9uhvAcde3"
       Case Is < 70000
         txtcodice = "N9uhvcde3"
         Text1.Text = "0polAk87y"
       Case Is < 74000
         txtcodice = "N0polk87yh"
         Text1.Text = "njhgAferr"
        Case Is < 78000
         txtcodice = "Nnjhgferrq"
         Text1.Text = "hfgjAur87"
          Case Is < 86000
         txtcodice = "Nhfgjur874"
         Text1.Text = "lkdrAfcvn"
       Case Is < 90000
         txtcodice = "Nlkdrfcvnm"
         Text1.Text = "rtyyAujdk"
        Case Is < 94000
         txtcodice = "Nrtyyujdkwj"
         Text1.Text = "eireAiofh"
       Case Is < 97000
         txtcodice = "Neireiofhf"
         Text1.Text = "hfaoAehhf"
       Case Is < 99000
         txtcodice = "Nhfaoehhfii"
         Text1.Text = "fjojArrij"
        Case Is >= 99000
         txtcodice = "Nfjojrrijgjr"
         Text1.Text = "1qscAvhu8"
       Case Else
         txtcodice = "Nhfaoehhfii"
         Text1.Text = "fjojArrij"
        
         End Select
         
         
          txtcodice = Right(SerialNumber, 2) + Text3 + txtcodice + Left(Text2, 5)
        
       
    
          Text1.Text = Right(SerialNumber, 2) + Left(Text2, 1) + Text1.Text + Text3
       
         
        
      
End Sub

Public Sub lingue()

Select Case lingua
  
 Case Is = "2I"
 messaggi(0) = "Chiedete il codice di attivazione"
messaggi(1) = "Login"
messaggi(2) = "Versione Demo"
messaggi(3) = "Password non valida"

   Exit Sub
   
   Case Is = "1F"
lblLabels(2).Caption = "Licence N."
lblLabels(0).Caption = "Code d'Installation"

lblLabels(1).Caption = "Code d'Activation"
cmdOK.Caption = "Valider"
command2.Caption = "Envoyer"
Command1.Caption = "Version de Demo"
cmdcancel.Caption = "Quitter"
Label2.Caption = "Pour recevoir le code d'activation envoyer le code d'installation et Votre Nom à l'e-mail support@erreasoft.com"
frmLogin.Caption = "Envoyer code d'installation à support@erreasoft.com"

messaggi(0) = "Vous devez demander the code d'activation"
messaggi(1) = "Accés"
messaggi(2) = "Version de demo. Vous pouvez insérer 5 clients en éprueve"
messaggi(3) = "Votre mot de passe il n'est pas valide"



Case Is = "3G"
lblLabels(2).Caption = "Licence No."
lblLabels(0).Caption = "Installation Code"
command2.Caption = "Send"
lblLabels(1).Caption = "Activation Code"
cmdOK.Caption = "Save"
Command1.Caption = "Demo Version"
cmdcancel.Caption = "Exit"
Label2.Caption = "To receive the activation code send Installation Code to e-mail support@erreasoft.com"
frmLogin.Caption = "Send Installation code to support@erreasoft.com"

messaggi(0) = "You can ask the activation code"
messaggi(1) = "Login"
messaggi(2) = "Demo Version"
messaggi(3) = "Invalid Password"

Case Is = "4S"
lblLabels(2).Caption = "Licenzia N."
lblLabels(0).Caption = "Codigo d'Installacion."
lblLabels(1).Caption = "Còdigo de l'activacion"
cmdOK.Caption = "Guarda"
command2.Caption = "Send"
Command1.Caption = "Demo Version"
cmdcancel.Caption = "Salir"
Label2.Caption = "Enviar codigo de l'installacion al email support@erreasoft.com para obtenir el código de la activacion. "
frmLogin.Caption = "Enviar codigo d'installacion al support@erreasoft.com"

messaggi(0) = "Usted puede preguntar el código de activación"
messaggi(1) = "Inicio de sesión"
messaggi(2) = "Versión de la demostración"
messaggi(3) = "Invalid Password"


Case Is = "5D"
lblLabels(2).Caption = "Licence No."
lblLabels(0).Caption = "Setupcode"
lblLabels(1).Caption = "Aktivierungscode"
cmdOK.Caption = "Übernehmen"
Command1.Caption = "Demo Version"
command2.Caption = "Send"
cmdcancel.Caption = "Beenden"
Label2.Caption = "Den Aktivierungscode zu erhalten, schicken Sie Genehmigungsnr. support@erreasoft.com zu E-mail"
 
frmLogin.Caption = "Schicken Sie die Genehmigungsnr. support@erreasoft.com zur E-mail"

messaggi(0) = "Sie können den Aktivierungscode fragen"
messaggi(1) = "Login"
messaggi(2) = "Demo Version"
messaggi(3) = "Invalid Password"

Case Else
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)

End

    
End Sub
'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    'initialise the strings
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    'call the API function
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
    
End Function
