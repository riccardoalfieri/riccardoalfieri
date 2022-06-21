VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmArtico 
   BackColor       =   &H00808080&
   Caption         =   "Produits"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14625
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   14625
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10695
      Left            =   -120
      TabIndex        =   24
      Top             =   0
      Width           =   14535
      Begin VB.CommandButton Command7 
         BackColor       =   &H0000FFFF&
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         MaskColor       =   &H0080FFFF&
         Picture         =   "frmARTICO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   8640
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   9480
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
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
         RecordSource    =   "SELECT * FROM tires"
         Caption         =   "AdoProduit"
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
      Begin VB.TextBox txtpic 
         DataField       =   "Immagine"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   9720
         TabIndex        =   68
         Text            =   "txtpic"
         Top             =   840
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0000FFFF&
         Caption         =   "Nouvelle Recherche"
         Height          =   495
         Left            =   3240
         Picture         =   "frmARTICO.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   7080
         Width           =   2295
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H008080FF&
         Caption         =   "Mémorise Image"
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   7080
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "...."
         Height          =   735
         Left            =   5760
         Picture         =   "frmARTICO.frx":0A14
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   3720
         Width           =   615
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFFF80&
         Caption         =   "...."
         Height          =   615
         Left            =   2640
         Picture         =   "frmARTICO.frx":2726
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   6720
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "CodiceInterno"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         DataField       =   "descrizione"
         DataSource      =   "Adodc1"
         Height          =   390
         Index           =   2
         Left            =   2400
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1425
         Width           =   5535
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "Giacenza"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   3
         Left            =   2400
         TabIndex        =   19
         Top             =   5760
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "listino1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   4
         Left            =   2400
         TabIndex        =   3
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "DataUltimaVendita"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   5
         Left            =   2640
         TabIndex        =   11
         Top             =   3345
         Width           =   3135
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "DataUltimoCarico"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   6
         Left            =   2640
         TabIndex        =   12
         Top             =   3660
         Width           =   3135
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   4440
         TabIndex        =   18
         Top             =   5400
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "PrezzoAcquisto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   10
         Left            =   2400
         TabIndex        =   15
         Top             =   4680
         Width           =   2895
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "ScortaMinima"
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   11
         Left            =   2400
         TabIndex        =   17
         Top             =   5400
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   39
         Top             =   6720
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "listino2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   7
         Left            =   2400
         TabIndex        =   5
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "per2"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   12
         Left            =   4920
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "per1"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   9
         Left            =   4920
         TabIndex        =   4
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "per3"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   13
         Left            =   4920
         TabIndex        =   8
         Top             =   2640
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "listino3"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   14
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         DataField       =   "PrezzoAcquistoNetto"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1040
            SubFormatType   =   1
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   285
         Index           =   15
         Left            =   2400
         TabIndex        =   16
         Top             =   5040
         Width           =   2895
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   2400
         TabIndex        =   9
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         DragMode        =   1  'Automatic
         Height          =   1875
         Left            =   8640
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   2  'Automatic
         Picture         =   "frmARTICO.frx":4438
         ScaleHeight     =   1815
         ScaleWidth      =   1530
         TabIndex        =   38
         Top             =   1800
         Width           =   1590
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   5640
         TabIndex        =   37
         Top             =   4440
         Width           =   2535
      End
      Begin VB.DirListBox Dir1 
         Height          =   1215
         Left            =   5640
         TabIndex        =   36
         Top             =   4680
         Width           =   2535
      End
      Begin VB.FileListBox File1 
         Height          =   1260
         Left            =   5640
         TabIndex        =   35
         Top             =   5880
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         DataField       =   "Itemdesc"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   5040
         TabIndex        =   10
         Top             =   3000
         Width           =   2535
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Height          =   255
         Left            =   2400
         TabIndex        =   22
         Top             =   6480
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Check3"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   6840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Cmdsearch 
         BackColor       =   &H0080FF80&
         Caption         =   "Rechercher"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6360
         Picture         =   "frmARTICO.frx":50B6
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   7680
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Nouveau"
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
         Left            =   960
         Picture         =   "frmARTICO.frx":6DC8
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FF8080&
         Caption         =   "Ajourner"
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
         Left            =   960
         Picture         =   "frmARTICO.frx":91BA
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   8640
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Valider"
         Enabled         =   0   'False
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
         Left            =   2760
         Picture         =   "frmARTICO.frx":9304
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   7680
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H008080FF&
         Caption         =   "Annuller"
         Enabled         =   0   'False
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
         Left            =   4560
         Picture         =   "frmARTICO.frx":944E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7680
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000000FF&
         Caption         =   "Supprimér"
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
         Left            =   2760
         Picture         =   "frmARTICO.frx":9890
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   8640
         Width           =   1695
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Quitter"
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
         Left            =   4560
         Picture         =   "frmARTICO.frx":99DA
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   8640
         Width           =   1575
      End
      Begin VB.TextBox TXTCODE 
         DataField       =   "CodiceInterno"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   7680
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo DataCombo4 
         Bindings        =   "frmARTICO.frx":9E1C
         DataField       =   "categorie"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2400
         TabIndex        =   21
         Top             =   6120
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Submenudesc"
         BoundColumn     =   "Submenudesc"
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "frmARTICO.frx":9E46
         DataField       =   "fornitore"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Top             =   3960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Fornitore"
         BoundColumn     =   "CodiceFornitore"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   5880
         Top             =   3840
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
         RecordSource    =   "SELECT * FROM Fornitori ORDER BY Fornitore"
         Caption         =   "Adodc1"
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
         Height          =   330
         Left            =   8280
         Top             =   4440
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
         RecordSource    =   "Patterns"
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
      Begin MSDataListLib.DataCombo DataCombo2 
         Bindings        =   "frmARTICO.frx":9E5B
         DataField       =   "misura"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   2400
         TabIndex        =   14
         Top             =   4320
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Pattern"
         Text            =   "PZ"
      End
      Begin MSDataListLib.DataCombo DataCombo3 
         Bindings        =   "frmARTICO.frx":9E70
         DataField       =   "iva"
         DataSource      =   "Adodc1"
         Height          =   315
         Left            =   4440
         TabIndex        =   20
         Top             =   5760
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "Size"
         Text            =   ""
      End
      Begin MSAdodcLib.Adodc Adodc4 
         Height          =   375
         Left            =   8280
         Top             =   5760
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
         RecordSource    =   "Sizes"
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
      Begin MSAdodcLib.Adodc Adosubmenu 
         Height          =   375
         Left            =   8280
         Top             =   6120
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
         RecordSource    =   "SELECT distinct( categorie) FROM tires "
         Caption         =   "AdoSubmenu"
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
      Begin MSComDlg.CommonDialog cdl 
         Left            =   120
         Top             =   6360
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmARTICO.frx":9E85
         Height          =   495
         Left            =   240
         TabIndex        =   26
         Top             =   9480
         Visible         =   0   'False
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   873
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   0   'False
         HeadLines       =   1
         RowHeight       =   19
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
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
            DataField       =   ""
            Caption         =   ""
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
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   855
         Left            =   720
         Picture         =   "frmARTICO.frx":9E9A
         Top             =   0
         Width           =   810
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Designation"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   63
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Codes "
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   62
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Qté en Stock"
         Height          =   255
         Index           =   3
         Left            =   480
         TabIndex        =   61
         Top             =   5760
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vente 1 TTC"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   60
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date Dernière Sortie"
         Height          =   255
         Index           =   5
         Left            =   480
         TabIndex        =   59
         Top             =   3345
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Date Dernière Entrée"
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   58
         Top             =   3660
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fournisseur"
         Height          =   255
         Index           =   7
         Left            =   480
         TabIndex        =   57
         Top             =   3975
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Mini"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   56
         Top             =   5400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Misure"
         Height          =   255
         Index           =   9
         Left            =   480
         TabIndex        =   55
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prix d'Achat "
         Height          =   255
         Index           =   10
         Left            =   480
         TabIndex        =   54
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Stock Maxi"
         Height          =   255
         Index           =   11
         Left            =   480
         TabIndex        =   53
         Top             =   5400
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rechercher codebarre"
         Height          =   255
         Left            =   3240
         TabIndex        =   52
         Top             =   6480
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vente 2 HTC"
         Height          =   255
         Index           =   12
         Left            =   480
         TabIndex        =   51
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TVA"
         Height          =   255
         Index           =   13
         Left            =   3720
         TabIndex        =   50
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coeff.pur ottenir vente 2"
         Height          =   255
         Index           =   14
         Left            =   3120
         TabIndex        =   49
         Top             =   2280
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coeff.pur ottenir vente 1"
         Height          =   255
         Index           =   15
         Left            =   3120
         TabIndex        =   48
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Coeff.pur ottenir vente 3"
         Height          =   255
         Index           =   16
         Left            =   3120
         TabIndex        =   47
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vente 3 HTC"
         Height          =   255
         Index           =   17
         Left            =   480
         TabIndex        =   46
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prix d'Achat Escompté"
         Height          =   255
         Index           =   18
         Left            =   480
         TabIndex        =   45
         Top             =   5040
         Width           =   2055
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Service"
         Height          =   255
         Index           =   19
         Left            =   480
         TabIndex        =   44
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categorie"
         Height          =   255
         Index           =   20
         Left            =   480
         TabIndex        =   43
         Top             =   6120
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ItemDesc"
         Height          =   255
         Index           =   21
         Left            =   3240
         TabIndex        =   42
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Visible"
         Height          =   255
         Index           =   22
         Left            =   600
         TabIndex        =   41
         Top             =   6480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formula"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   600
         TabIndex        =   40
         Top             =   6840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliquez Nouveau, Insérer le Produit et Cliquez Valider"
         Height          =   375
         Left            =   1680
         TabIndex        =   34
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Produits"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   1680
         TabIndex        =   33
         Top             =   120
         Width           =   4335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         Height          =   1095
         Left            =   360
         Top             =   -120
         Width           =   9615
      End
   End
End
Attribute VB_Name = "frmArtico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mb



Private Sub cmdAdd_Click()
Adodc1.Recordset.AddNew

addupdate

Randomize
TXTCODE = "TIR" & Round(Rnd() * 999999) & TXTCODE + Chr(Round(Rnd() * 25) + 65)
End Sub

Private Sub cmdBack_Click()

Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmArtico.Show
End Sub

Private Sub cmdDelete_Click()
Dim response, stringa As String
stringa = TXTCODE & ".jpg"

On Error Resume Next
response = MsgBox("Supprimér ce produit?", vbOKCancel + vbCancel, "Attention")
Select Case response
 Case 6
 Adodc1.Recordset.delete
 Kill stringa

 Case Else
End Select





End Sub

Private Sub cmdReport_Click()


End Sub

Private Sub cmdSave_Click()

On Error Resume Next
If Len(Text3) = 0 Then
Text3 = Left(txtFields(2), 15)
End If
If Len(txtFields(2)) > 0 Then
    Adodc1.Recordset.UpdateBatch
    savecancel
    delete
Else
    mb = MsgBox("Inserire articolo ", vbCritical, "Attention")
    txtFields(2).SetFocus
End If

End Sub



Private Sub Cmdsearch_Click()
Dim stringa As String
Picture1.Picture = Nothing
Form6.Show vbModal
On Error Resume Next
              Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "[codiceinterno] = '" & art1 & "'", 0
 stringa = TXTCODE & ".jpg"
 Picture1.Picture = LoadPicture(stringa)

End Sub

Private Sub cmdUpdate_Click()
addupdate
End Sub







Private Sub Command2_Click()
Text1.Text = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
Dim stringa As String
Picture1.Picture = Nothing
On Error Resume Next
frmricerca.Show vbModal
        Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "codiceinterno = '" & art1 & "'", 0

 stringa = TXTCODE & ".jpg"
 Picture1.Picture = LoadPicture(stringa)


End Sub

Private Sub Command4_Click()
On Error GoTo dopo
Dim stringa As String
stringa = TXTCODE & ".jpg"

SavePicture Picture1.Picture, stringa
txtpic = stringa
dopo:
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
'rtb.Text = ""
Picture1.Cls

File1.Pattern = "*.jpg;*.bmp;*.gif"
End Sub

Private Sub Drive1_Change()
On Error GoTo errhnd
Dir1.Path = Drive1.Drive
errhnd:
Select Case Err.Number
Dim msg As String
Case 68
msg = "your this drive are not resource" & vbNewLine
msg = msg + "1. try or look your cdrom or floppy disk drive" & vbNewLine
msg = msg + " 2. resource are not available " & vbNewLine
msg = msg + " 3. we are set drive default is c:\"
MsgBox msg, vbOKOnly + vbExclamation, "please check"
 Drive1.Drive = "C:\"
 End Select
End Sub

Private Sub File1_Click()
On Error Resume Next
Picture1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)

End Sub

Private Sub Form_Activate()
Dim I As Integer

File1.Pattern = "*.jpg;*.bmp;*.gif"


delete

End Sub

Private Function addupdate()
cmdAdd.Enabled = False
cmdUpdate.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
cmdDelete.Enabled = False
'cmdReport.Enabled = False

txtFields(2).Locked = False
txtFields(3).Locked = False
'txtNumber.Locked = False
'txtName.SetFocus


Adodc1.Enabled = False
DataGrid1.Enabled = False
End Function

Private Function savecancel()
DataGrid1.Refresh

cmdAdd.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False

txtFields(2).Locked = True
txtFields(3).Locked = True
'txtNumber.Locked = True
End Function

Private Function delete()
DataGrid1.Refresh

If Adodc1.Recordset.RecordCount = 0 Then
    Adodc1.Enabled = False
    DataGrid1.Enabled = False
    
    cmdDelete.Enabled = False
    cmdUpdate.Enabled = False
 '   cmdReport.Enabled = False
Else
    Adodc1.Enabled = True
    DataGrid1.Enabled = True
    
    cmdDelete.Enabled = True
    cmdUpdate.Enabled = True
   ' cmdReport.Enabled = True
End If

End Function

Private Sub Form_Load()
'MonthView1.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

Adodc1.Recordset.CancelBatch
End Sub

Private Sub Frame1_Click()
'MonthView1.Visible = False
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
'txtanniv = MonthView1
'MonthView1.Visible = False

End Sub

Private Sub txtanniv_Click()
'MonthView1.Visible = True
'MonthView1.Value = Date

End Sub


Private Sub Text1_KeyDown( _
           KeyCode As Integer, Shift As Integer)
    Dim stringa As String
On Error Resume Next

     Select Case KeyCode
     Case vbKeyReturn:
               Adodc1.Recordset.MoveFirst
  Adodc1.Recordset.Find "codiceean = '" & Text1.Text & "'", 0

 stringa = TXTCODE & ".jpg"
 Picture1.Picture = LoadPicture(stringa)



     End Select

End Sub

Private Sub txtFields_Change(Index As Integer)
Dim stringa As String
On Error Resume Next
Select Case Index
         Case 0, 1, 2
          stringa = TXTCODE & ".jpg"
 Picture1.Picture = LoadPicture(Dir1.Path & "\" & stringa)
  Case Else
  End Select
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

     Select Case KeyCode
     Case vbKeyReturn:
              
              


      Case Else
     End Select

End Sub
