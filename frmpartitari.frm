VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmpartitari 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   Caption         =   "Vendors Account"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13620
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13620
   Begin VB.CommandButton cmdback 
      BackColor       =   &H0080FF80&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10320
      Picture         =   "frmpartitari.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7440
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0000FFFF&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8760
      MaskColor       =   &H0080FFFF&
      Picture         =   "frmpartitari.frx":1D42
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7440
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6840
      Picture         =   "frmpartitari.frx":260C
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7440
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Update"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5040
      Picture         =   "frmpartitari.frx":2756
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7440
      Width           =   1815
   End
   Begin VB.CommandButton command3 
      BackColor       =   &H00FF8080&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      Picture         =   "frmpartitari.frx":4450
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7440
      Width           =   2055
   End
   Begin VB.CommandButton cmdstampa 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      Picture         =   "frmpartitari.frx":6842
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   735
      Left            =   3720
      Picture         =   "frmpartitari.frx":8554
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6960
      TabIndex        =   20
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5400
      TabIndex        =   19
      Top             =   4800
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Top             =   4800
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "frmpartitari.frx":A266
      Height          =   1335
      Left            =   6120
      TabIndex        =   17
      Top             =   5760
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   2355
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      Caption         =   "Rate"
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "codicefornitore"
         Caption         =   "codicefornitore"
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
         DataField       =   "fornitore"
         Caption         =   "fornitore"
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
         DataField       =   "numerofattura"
         Caption         =   "numerofattura"
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
         DataField       =   "datafattura"
         Caption         =   "datafattura"
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
         DataField       =   "datascadenza"
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
      BeginProperty Column06 
         DataField       =   "importopagamento"
         Caption         =   "Amount"
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
         DataField       =   "pagato"
         Caption         =   "pagato"
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
         DataField       =   "libero"
         Caption         =   "libero"
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "frmpartitari.frx":A27B
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   5760
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1931
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
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "codice"
         Caption         =   "codice"
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
         DataField       =   "tipoPagamento"
         Caption         =   "Payment"
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
         DataField       =   "scadenza30"
         Caption         =   "30 "
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
         DataField       =   "scadenza60"
         Caption         =   "60 "
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
         DataField       =   "scadenza90"
         Caption         =   "90 "
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
         DataField       =   "scadenza120"
         Caption         =   "120 "
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
         DataField       =   "scadenza150"
         Caption         =   "150 "
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
         DataField       =   "scadenza180"
         Caption         =   "180 "
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
         DataField       =   "numerorate"
         Caption         =   "numerorate"
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
         DataField       =   "finemese"
         Caption         =   "finemese"
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
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtnumerobis 
      DataField       =   "numerorate"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   0
      TabIndex        =   15
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmpartitari.frx":A290
      Height          =   315
      Left            =   7920
      TabIndex        =   14
      Top             =   5400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "tipoPagamento"
      Text            =   ""
   End
   Begin VB.TextBox txtdtadoc 
      Height          =   360
      Left            =   6960
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.TextBox txtndoc 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Text            =   "0"
      Top             =   5400
      Width           =   1455
   End
   Begin VB.TextBox txtdata 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   5400
      Width           =   975
   End
   Begin VB.ComboBox cbcausale 
      Height          =   315
      Left            =   1080
      TabIndex        =   1
      Top             =   5400
      Width           =   3135
   End
   Begin VB.TextBox txtimporto 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   5400
      Width           =   1335
   End
   Begin VB.TextBox txtSupplier 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "frmpartitari.frx":A2A5
      Height          =   4215
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   18615
      _ExtentX        =   32835
      _ExtentY        =   7435
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "List of Movements"
      ColumnCount     =   24
      BeginProperty Column00 
         DataField       =   "Data"
         Caption         =   "Date"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Causale"
         Caption         =   "Type"
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
         DataField       =   "ImportoDare"
         Caption         =   "Payement"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ImportoAvere"
         Caption         =   "Invoice"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "NumeroDocumento"
         Caption         =   "No."
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
         DataField       =   "DataDocumento"
         Caption         =   "Invoice Date"
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
         DataField       =   "TipoPagamento"
         Caption         =   "Payment"
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
         DataField       =   "DataPagamento"
         Caption         =   "Date "
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
         DataField       =   "Codice"
         Caption         =   "Code"
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
         DataField       =   "Fornitore"
         Caption         =   "Vendor"
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
         DataField       =   "SaldoProgressivo"
         Caption         =   "SaldoProgressivo"
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
         DataField       =   "Commento"
         Caption         =   "Commento"
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
         DataField       =   "data30"
         Caption         =   "data30"
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
         DataField       =   "importo30"
         Caption         =   "importo30"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "data60"
         Caption         =   "data60"
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
         DataField       =   "importo60"
         Caption         =   "importo60"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column16 
         DataField       =   "data90"
         Caption         =   "data90"
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
      BeginProperty Column17 
         DataField       =   "importo90"
         Caption         =   "importo90"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column18 
         DataField       =   "data120"
         Caption         =   "data120"
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
      BeginProperty Column19 
         DataField       =   "importo120"
         Caption         =   "importo120"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column20 
         DataField       =   "data150"
         Caption         =   "data150"
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
      BeginProperty Column21 
         DataField       =   "importo150"
         Caption         =   "importo150"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column22 
         DataField       =   "data180"
         Caption         =   "data180"
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
      BeginProperty Column23 
         DataField       =   "importo180"
         Caption         =   "importo180"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
         EndProperty
         BeginProperty Column10 
         EndProperty
         BeginProperty Column11 
         EndProperty
         BeginProperty Column12 
         EndProperty
         BeginProperty Column13 
            Alignment       =   1
         EndProperty
         BeginProperty Column14 
         EndProperty
         BeginProperty Column15 
            Alignment       =   1
         EndProperty
         BeginProperty Column16 
         EndProperty
         BeginProperty Column17 
         EndProperty
         BeginProperty Column18 
         EndProperty
         BeginProperty Column19 
            Alignment       =   1
         EndProperty
         BeginProperty Column20 
         EndProperty
         BeginProperty Column21 
            Alignment       =   1
         EndProperty
         BeginProperty Column22 
         EndProperty
         BeginProperty Column23 
            Alignment       =   1
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   855
      Left            =   7080
      Top             =   240
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1508
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
      RecordSource    =   "select * from Purchase where causale="""""
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Left            =   9720
      Top             =   5040
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      RecordSource    =   "select * from archiviopagamenti"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Left            =   1680
      Top             =   8160
      Visible         =   0   'False
      Width           =   1680
      _ExtentX        =   2963
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
      RecordSource    =   "select * from archiviopagamenti"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      Height          =   375
      Left            =   5520
      Top             =   8400
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      RecordSource    =   "select * from scadenzario where codicefornitore="""""
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inv.Date"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No"
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Payment Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1080
      TabIndex        =   11
      Top             =   5160
      Width           =   3165
   End
   Begin VB.Label lblpezzi 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   1095
   End
   Begin VB.Label lblimporto 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lblpagamento 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Payment"
      Height          =   255
      Left            =   7920
      TabIndex        =   8
      Top             =   5160
      Width           =   3495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Vendor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   0
      Width           =   690
   End
End
Attribute VB_Name = "frmpartitari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cbcausale_Change()

Select Case cbcausale
  Case "Fattura di Acquisto", "Nota di Credito Fornitore"
    txtndoc.Enabled = True
    txtdtadoc.Enabled = True
   Case Else
    txtndoc.Enabled = False
     txtdtadoc.Enabled = False
    End Select
End Sub

Private Sub cmdBack_Click()

Unload Me

End Sub

Private Sub cmdstampa_Click()

Set DataReport303.DataSource = Adodc5

With DataReport303
With .Sections("Section4")
           With .Controls("label2")
            .Caption = "Conto Fornitori"
            End With
            With .Controls("label12")
            .Caption = txtSupplier
            End With
            
        End With
        
With .Sections("Section2")
            
            With .Controls
             .Item("label7").Caption = "Data Doc."
             .Item("label1").Caption = "Documento"
             .Item("label4").Caption = "Numero Doc."
             .Item("label8").Caption = "Imp.Avere"
             .Item("label3").Caption = "Imp.Dare"
          
             End With
        End With
        
       With .Sections("Section5")
            
            With .Controls
             .Item("label11").Caption = "Tot.Avere"
             .Item("label6").Caption = Text1
             .Item("label9").Caption = "Tot.Dare"
             .Item("label10").Caption = Text2
             .Item("label13").Caption = "Saldo"
           .Item("label14").Caption = Text3
             End With
        End With
    .Show
  End With


End Sub

Private Sub Command1_Click()

Form7.Show vbModal

cbcausale.Enabled = True
aggiorna

 
End Sub

Private Sub Command2_Click()
If txtimporto = "" Then GoTo dopo
If CSng(txtimporto) = 0 Then GoTo dopo

Select Case cbcausale
  Case "Fattura di Acquisto"
  DataGrid5.Columns(3) = txtimporto
  DataGrid5.Columns(2) = ""
     
     scadenze
     
   Case Else
     DataGrid5.Columns(2) = txtimporto
     DataGrid5.Columns(3) = ""
    End Select
    
    DataGrid5.Columns(0) = txtdata
   
    DataGrid5.Columns(1) = cbcausale
    DataGrid5.Columns(11) = txtnote
    DataGrid5.Columns(4) = txtndoc
    DataGrid5.Columns(5) = txtdtadoc
    DataGrid5.Columns(6) = DataCombo1
    DataGrid5.Columns(7) = txtdatapag
    DataGrid5.Columns(8) = variabile1
     DataGrid5.Columns(9) = txtSupplier
dopo:
txtnote = ""
txtndoc = ""
txtdtadoc = ""
DataCombo1 = ""
txtdatapag = ""
cbcausale = ""
txtimporto = ""
'Command3.Enabled = True
'Command4.Enabled = False
'Command2.Enabled = False




End Sub

Private Sub Command3_Click()
'On Error Resume Next
If txtimporto = "" Then GoTo dopo
  If CSng(txtimporto) = 0 Then GoTo dopo
    If txtdoc = "" Then txtdoc = "1"
     If txtdtadoc = "" Then txtdtadoc = Date
     If txtdata = "" Then txtdata = Date
     
Adodc5.Recordset.AddNew
Select Case cbcausale.ListIndex
  Case 0
   DataGrid5.Columns(3) = txtimporto
    
  scadenze
                     
             
            
   Case Else
    DataGrid5.Columns(2) = txtimporto
    End Select
    
    DataGrid5.Columns(0) = txtdata
   
    DataGrid5.Columns(1) = cbcausale
    DataGrid5.Columns(11) = txtnote
    DataGrid5.Columns(4) = txtndoc
    DataGrid5.Columns(5) = txtdtadoc
    DataGrid5.Columns(6) = DataCombo1
    DataGrid5.Columns(7) = txtdatapag
    DataGrid5.Columns(8) = variabile1
     DataGrid5.Columns(9) = txtSupplier
     
    
     
    
     
     

             
     Adodc5.Recordset.Update
dopo:
txtnote = ""
txtndoc = ""
txtdtadoc = ""
DataCombo1 = ""
txtdatapag = ""
cbcausale = ""
txtimporto = ""
DataCombo1 = ""
Command1.SetFocus
'Command3.Enabled = False

aggiorna
End Sub

Private Sub Command3_LostFocus()
aggiorna
End Sub

Private Sub Command4_Click()
txtnote = ""
txtndoc = ""
txtdtadoc = ""
DataCombo1 = ""
txtdatapag = ""
cbcausale = ""
txtimporto = ""
'Command3.Enabled = True
'Command2.Enabled = False
'Command4.Enabled = False
End Sub

Private Sub Command5_Click()
On Error Resume Next
response = MsgBox(messaggi(9), vbOKCancel + vbCancel, messaggi(8))
Select Case response
 Case 6
 Adodc5.Recordset.delete
 
 Case Else
 
 End Select
Command5_LostFocus

Command1.SetFocus

End Sub

Private Sub Command5_LostFocus()
aggiorna
End Sub

Private Sub DataCombo1_Change()
Dim stringa As String
stringa = "select * from " & archivipagamenti & " where tipopagamento= '" & DataCombo1 & "'"

  With Adodc3
     .RecordSource = stringa
     .Refresh
    End With
    
    
    With DataGrid2
     .ClearFields
     .HoldFields
     .ReBind
    
     End With
End Sub


Private Sub DataGrid5_Click()
Dim stringa As String
On Error Resume Next



'Command3.Enabled = False
'Command2.Enabled = True
'Command4.Enabled = True
'Command5.Enabled = True
txtnote = DataGrid5.Columns(11)
txtndoc = DataGrid5.Columns(4)
txtdtadoc = DataGrid5.Columns(5)
DataCombo1 = DataGrid5.Columns(6)
txtdatapag = DataGrid5.Columns(7)
cbcausale = DataGrid5.Columns(1)
Select Case cbcausale
  Case "Fattura di Acquisto"
  txtimporto = DataGrid5.Columns(3)
   Case Else
     txtimporto = DataGrid5.Columns(2)
    End Select

stringa = "select * from scadenzario where codicefornitore= '" & variabile1 & "' AND numerofattura = '" & DataGrid5.Columns(4) & "'"

  With Adodc4
     .RecordSource = stringa
     .Refresh
    End With
    
    
    With DataGrid3
     .ClearFields
     .HoldFields
     .ReBind
    
     End With




End Sub

Private Sub DataGrid5_DblClick()

'Command3.Enabled = False
End Sub

Private Sub Form_Load()
lingue

Dim stringa As String
stringa = "select * from " & archivipagamenti

  With Adodc2
     .RecordSource = stringa
     .Refresh
    End With
    
    
    With DataCombo1
  
    
     .Refresh
    
     End With

txtdata = Date

'cbcausale.Enabled = False
'Command2.Enabled = False
'Command3.Enabled = True
'Command4.Enabled = False
'Command5.Enabled = False
End Sub



Private Sub txtdatapag_Validate(Cancel As Boolean)
  ' Prepare to edit in short-date format.
    On Error Resume Next
    txtdatapag.Text = Format$(CDate(txtdatapag.Text), "short date")
End Sub

Private Sub txtdtadoc_Validate(Cancel As Boolean)
  ' Prepare to edit in short-date format.
    On Error Resume Next
    txtdtadoc.Text = Format$(CDate(txtdtadoc.Text), "short date")
    m = Month(txtdtadoc)
 
End Sub

Public Sub aggiorna()

Dim stringa As String
 stringa = "SELECT * FROM Purchase WHERE codice = '" & variabile1 & "' order by datadocumento, cint(NumeroDocumento)"

   With Adodc5
    .RecordSource = stringa
    .Refresh
    End With
    
    
    With DataGrid5
     .ClearFields
     .HoldFields
     .ReBind
    
     End With



 On Error Resume Next
a = 0: b = 0: c = 0
Adodc5.Recordset.MoveFirst
 For I = 0 To Adodc5.Recordset.RecordCount - 1
  a = a + DataGrid5.Columns(2)
   b = b + DataGrid5.Columns(3)
    c = a - b
     Adodc5.Recordset.MoveNext
   Next I
   
     Text1 = Format((a), "###,###.00")
    Text2 = Format((b), "###,###.00")
     Text3 = Format((c), "###,###.00")

txtSupplier = variabile2

 If Adodc5.Recordset.RecordCount > 0 Then
 DataGrid5.Enabled = True
 Else
 DataGrid5.Enabled = False
 End If
 
End Sub


Public Sub scadenze()
On Error GoTo fine
  'CALCOLA SCADENZE
   
       txtnumerobis = (txtimporto) / (DataGrid2.Columns(9))
 Y = Year(txtdtadoc)
 m = Month(txtdtadoc)
 
 
      For I = 0 To 5
       If DataGrid2.Columns(3 + I).Value <> 0 Then
       Adodc4.Recordset.AddNew
         DataGrid3.Columns(1) = variabile1 ' codice fornitore
          DataGrid3.Columns(2) = txtSupplier '  fornitore
          DataGrid3.Columns(3) = txtndoc
          DataGrid3.Columns(4) = txtdtadoc
          
             DataGrid3.Columns(6) = txtnumerobis ' importo scadenza
             
             Select Case (m + I + 1)
                Case Is > 12
                mese% = (m + I + 1) - 12: newy = Y + 1
                 
                  Case Else
                  mese% = m + I + 1: newy = Y
                  End Select
                  
                  
                     DataGrid3.Columns(5) = DateSerial(newy, mese%, 30) ' data scadenza
             
             Else
             
             End If
              Next I
              
fine:
              
End Sub
Public Sub lingue()

Select Case lingua
  
 Case Is = "1F"
 Label2 = "Fournisseurs"
DataGrid5.Caption = " Compte Fournisseurs"
DataGrid5.Columns(0).Caption = "Date"
DataGrid5.Columns(1).Caption = "Nature"
DataGrid5.Columns(2).Caption = "Remise"
DataGrid5.Columns(3).Caption = "Facture"
DataGrid5.Columns(4).Caption = "Doc.N."
DataGrid5.Columns(5).Caption = "Doc.Date"
DataGrid5.Columns(6).Caption = "Paiement"
DataGrid5.Columns(7).Caption = "Date Paiement"
DataGrid5.Columns(8).Caption = "Code"
DataGrid5.Columns(9).Caption = "Client"
DataGrid5.Columns(10).Caption = "Total"
lblpezzi = "Date"
Label13 = "Nature"
lblimporto = "Total"
Label1 = "N.Doc."
Label3 = "Date Doc."
lblpagamento = "Paiement"
DataGrid2.Columns(2).Caption = "Paiement"
DataGrid2.Columns(3).Caption = "30 jours"
DataGrid2.Columns(4).Caption = "60 jours"
DataGrid2.Columns(5).Caption = "90 jours"
DataGrid2.Columns(6).Caption = "120 jours"
DataGrid2.Columns(7).Caption = "150 jours"
DataGrid2.Columns(8).Caption = "180 jours"


DataGrid3.Caption = "Echéancies"
DataGrid3.Columns(5).Caption = "Date"
DataGrid3.Columns(6).Caption = "Echéances"

cmdstampa.Caption = "Imprimer"

Command4.Caption = "Nouveau"
Command3.Caption = "Valider"
Command5.Caption = "Supprimér"
Command2.Caption = "Ajourner"
cmdBack.Caption = "Quitter"

frmpartitari.Caption = "Compte Fournisseurs"
 cbcausale.AddItem "Facture"
cbcausale.AddItem "Crédit fournisseur"
cbcausale.AddItem "Paiement"
cbcausale.AddItem "Autres"

  archivipagamenti = "ArchivioPagamenti"

   Exit Sub
   
   Case Is = "2I"
Label2 = "Fornitori"
DataGrid5.Caption = " Conto Fornitori"
DataGrid5.Columns(0).Caption = "Data"
DataGrid5.Columns(1).Caption = "Causale"
DataGrid5.Columns(2).Caption = "Avere"
DataGrid5.Columns(3).Caption = "Dare"
DataGrid5.Columns(4).Caption = "Doc.N."
DataGrid5.Columns(5).Caption = "Doc.Data"
DataGrid5.Columns(6).Caption = "Tipo Pag."
DataGrid5.Columns(7).Caption = "Data Pag."
DataGrid5.Columns(8).Caption = "Codice"
DataGrid5.Columns(9).Caption = "Cliente"
DataGrid5.Columns(10).Caption = "Saldo"
lblpezzi = "Data"
Label13 = "Causale"
lblimporto = "Totale"
Label1 = "N.Doc."
Label3 = "Data Doc."
lblpagamento = "Pagamento"
DataGrid2.Columns(2).Caption = "Pagamento"
DataGrid2.Columns(3).Caption = "30 gg"
DataGrid2.Columns(4).Caption = "60 gg"
DataGrid2.Columns(5).Caption = "90 gg"
DataGrid2.Columns(6).Caption = "120 gg"
DataGrid2.Columns(7).Caption = "150 gg"
DataGrid2.Columns(8).Caption = "180 gg"


DataGrid3.Caption = "Scadenze Pagamenti"
DataGrid3.Columns(5).Caption = "Data"
DataGrid3.Columns(6).Caption = "Importo"

cmdstampa.Caption = "Stampa"

Command4.Caption = "Nuovo"
Command3.Caption = "Salva"
Command5.Caption = "Elimina"
Command2.Caption = "Aggiorna"
cmdBack.Caption = "Esci"

frmpartitari.Caption = "Conto Fonitori"
 cbcausale.AddItem "Fattura"
cbcausale.AddItem "Nota di Credito"
cbcausale.AddItem "Pagamento"
cbcausale.AddItem "Storno su Fattura"
cbcausale.AddItem "Altri"

  archivipagamenti = "ArchivioPagamentiIT"


 Case Is = "3G"
 
Label2 = "Vendor"
DataGrid5.Caption = "Vendors Account"
DataGrid5.Columns(0).Caption = "Date"
DataGrid5.Columns(1).Caption = "Document"
DataGrid5.Columns(2).Caption = "Invoice Amount"
DataGrid5.Columns(3).Caption = "Amount"
DataGrid5.Columns(4).Caption = "Doc.No."
DataGrid5.Columns(5).Caption = "Doc.Date"
DataGrid5.Columns(6).Caption = "Payment."
DataGrid5.Columns(7).Caption = "Payment Date"
DataGrid5.Columns(8).Caption = "Code"
DataGrid5.Columns(9).Caption = "Client"
DataGrid5.Columns(10).Caption = "Total"
lblpezzi = "Date"
Label13 = "Document"
lblimporto = "Total "
Label1 = "Doc.No."
Label3 = "Doc.Date"
lblpagamento = "Payment"
DataGrid2.Columns(2).Caption = "Payment"
DataGrid2.Columns(3).Caption = "30 days"
DataGrid2.Columns(4).Caption = "60 days"
DataGrid2.Columns(5).Caption = "90 days"
DataGrid2.Columns(6).Caption = "120 days"
DataGrid2.Columns(7).Caption = "150 days"
DataGrid2.Columns(8).Caption = "180 days"


DataGrid3.Caption = "Payments Terms"
DataGrid3.Columns(5).Caption = "Date"
DataGrid3.Columns(6).Caption = "Amount"

cmdstampa.Caption = "Print"

Command4.Caption = "New"
Command3.Caption = "Save"
Command5.Caption = "Delete"
Command2.Caption = "Update"
cmdBack.Caption = "Exit"

frmpartitari.Caption = "Vendors Account"

 cbcausale.AddItem "Invoice"
cbcausale.AddItem "Vendors Credit Invoice"
cbcausale.AddItem "Payment"
cbcausale.AddItem "Discount on Vendors Invoice"
cbcausale.AddItem "Others"
 
 archivipagamenti = "ArchivioPagamentiEN"
 
 Case Is = "4S"
 
Label2 = "Proveedor"
DataGrid5.Caption = "Pagos Facturas Proveedor"
DataGrid5.Columns(0).Caption = "Fecha"
DataGrid5.Columns(1).Caption = "Movimiento"
DataGrid5.Columns(2).Caption = "Importo"
DataGrid5.Columns(3).Caption = "Importo Factura"
DataGrid5.Columns(4).Caption = "Doc. N."
DataGrid5.Columns(5).Caption = "Fecha Doc."
DataGrid5.Columns(6).Caption = "Pago"
DataGrid5.Columns(7).Caption = "Fecha Pago"
DataGrid5.Columns(8).Caption = "Codigo"
DataGrid5.Columns(9).Caption = "Proveedor"
DataGrid5.Columns(10).Caption = "Total"
lblpezzi = "Fecha"
Label13 = "Document"
lblimporto = "Total "
Label1 = "n. Doc.."
Label3 = "Fecha Doc."
lblpagamento = "Pago"
DataGrid2.Columns(2).Caption = "Pago"
DataGrid2.Columns(3).Caption = "30 dias"
DataGrid2.Columns(4).Caption = "60 dias"
DataGrid2.Columns(5).Caption = "90 dias"
DataGrid2.Columns(6).Caption = "120 dias"
DataGrid2.Columns(7).Caption = "150 dias"
DataGrid2.Columns(8).Caption = "180 dias"


DataGrid3.Caption = "Terminos de Pago"
DataGrid3.Columns(5).Caption = "Fecha"
DataGrid3.Columns(6).Caption = "Total"

cmdstampa.Caption = "Impresión"

Command4.Caption = "Nuevo"
Command3.Caption = "Guarda"
Command5.Caption = "Borra"
Command2.Caption = "Actualiza"
cmdBack.Caption = "Salir"

frmpartitari.Caption = "Pagos Facturas Proveedor"

 cbcausale.AddItem "Factura"
cbcausale.AddItem "Nota Crédito Proveedor"
cbcausale.AddItem "Pago"
cbcausale.AddItem "Descuento"
cbcausale.AddItem "Otros"
 
 archivipagamenti = "ArchivioPagamentiSP"
 
 Case Is = "5D"
 
Label2 = "Lieferant"
DataGrid5.Caption = "Lieferantkonto"
DataGrid5.Columns(0).Caption = "Datum"
DataGrid5.Columns(1).Caption = "Dokument"
DataGrid5.Columns(2).Caption = "Menge"
DataGrid5.Columns(3).Caption = "Gesamte Rechnung"
DataGrid5.Columns(4).Caption = "Dok.No."
DataGrid5.Columns(5).Caption = "Dok.Datum"
DataGrid5.Columns(6).Caption = "Zahlung"
DataGrid5.Columns(7).Caption = "Zahlung Datum"
DataGrid5.Columns(8).Caption = "Code"
DataGrid5.Columns(9).Caption = "Kunde"
DataGrid5.Columns(10).Caption = "Gesamte"
lblpezzi = "Datum"
Label13 = "Dokument"
lblimporto = "Gesamte "
Label1 = "Dok.No."
Label3 = "Dok.Datum"
lblpagamento = "Zahlung"
DataGrid2.Columns(2).Caption = "Zahlung"
DataGrid2.Columns(3).Caption = "30 Tage"
DataGrid2.Columns(4).Caption = "60 Tage"
DataGrid2.Columns(5).Caption = "90 Tage"
DataGrid2.Columns(6).Caption = "120 Tage"
DataGrid2.Columns(7).Caption = "150 Tage"
DataGrid2.Columns(8).Caption = "180 Tage"


DataGrid3.Caption = "Zahlungstermine"
DataGrid3.Columns(5).Caption = "Datum"
DataGrid3.Columns(6).Caption = "Gesamte"

cmdstampa.Caption = "Drucke"

Command4.Caption = "Neu"
Command3.Caption = "Übernehmen"
Command5.Caption = "Löschen"
Command2.Caption = "Update"
cmdBack.Caption = "Beenden"


frmpartitari.Caption = "Lieferantkonto"

 cbcausale.AddItem "Rechnung"
cbcausale.AddItem "Kredit"
cbcausale.AddItem "Erhalten Zahlungen"
cbcausale.AddItem "Rabatt"
cbcausale.AddItem "Others"
 
 archivipagamenti = "ArchivioPagamentiDE"
 
Case Else
End Select
 
End Sub


