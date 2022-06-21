VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMARTICO 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produits"
   ClientHeight    =   9045
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   15270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   9045
   ScaleWidth      =   15270
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      DragMode        =   1  'Automatic
      Height          =   1875
      Left            =   8040
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      Picture         =   "frmARTICO1.frx":0000
      ScaleHeight     =   1815
      ScaleWidth      =   1530
      TabIndex        =   58
      Top             =   0
      Width           =   1590
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   5520
      TabIndex        =   57
      Top             =   4080
      Width           =   2535
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   5520
      TabIndex        =   56
      Top             =   4440
      Width           =   2535
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   5520
      TabIndex        =   55
      Top             =   5640
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Mémorise image"
      Height          =   495
      Left            =   5520
      TabIndex        =   54
      Top             =   6960
      Width           =   2535
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "codiceEAN"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "CodiceInterno"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   320
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "descrizione"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   2040
      TabIndex        =   2
      Top             =   640
      Width           =   3375
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   960
      Width           =   1095
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   4
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "DataUltimaVendita"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Top             =   2565
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "DataUltimoCarico"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   2880
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "LottoRiordino"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   8
      Left            =   2040
      TabIndex        =   10
      Top             =   3525
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "PrezzoAcquisto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   10
      Left            =   2040
      TabIndex        =   12
      Top             =   4260
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "ScortaMinima"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   11
      Left            =   2040
      TabIndex        =   14
      Top             =   4980
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3360
      TabIndex        =   30
      Top             =   5700
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Nouvelle Recherche"
      Height          =   495
      Left            =   3360
      Picture         =   "frmARTICO1.frx":0C7E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6060
      Width           =   2055
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   7
      Left            =   2040
      TabIndex        =   5
      Top             =   1620
      Visible         =   0   'False
      Width           =   1095
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   12
      Left            =   5160
      TabIndex        =   28
      Top             =   1620
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   9
      Left            =   5160
      TabIndex        =   27
      Top             =   1260
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   13
      Left            =   4440
      TabIndex        =   26
      Top             =   2220
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
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   14
      Left            =   6600
      TabIndex        =   25
      Top             =   1980
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "PrezzoAcquistoNetto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   15
      Left            =   2040
      TabIndex        =   13
      Top             =   4620
      Width           =   3375
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Quitter"
      Height          =   900
      Left            =   4680
      Picture         =   "frmARTICO1.frx":0DC8
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Ajourner"
      Height          =   900
      Left            =   3600
      Picture         =   "frmARTICO1.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Supprimer"
      Height          =   900
      Left            =   2520
      Picture         =   "frmARTICO1.frx":1354
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Valider"
      Height          =   900
      Left            =   1440
      Picture         =   "frmARTICO1.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7740
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Nouveau"
      Height          =   900
      Left            =   360
      Picture         =   "frmARTICO1.frx":15E8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7740
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      DataField       =   "reparto"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1040
         SubFormatType   =   1
      EndProperty
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   16
      Left            =   5160
      TabIndex        =   19
      Top             =   940
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "...."
      Height          =   735
      Left            =   5400
      Picture         =   "frmARTICO1.frx":39DA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "...."
      Height          =   735
      Left            =   5400
      Picture         =   "frmARTICO1.frx":56EC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2940
      Width           =   615
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      DataField       =   "service"
      DataSource      =   "datPrimaryRS"
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   2220
      Width           =   255
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frmARTICO1.frx":73FE
      DataField       =   "Fornitore"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   9
      Top             =   3180
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Fornitore"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc datPrimaryRS 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   8715
      Width           =   15270
      _ExtentX        =   26935
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
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from tires"
      Caption         =   " "
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   5760
      Top             =   3180
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   5520
      Top             =   3900
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
      Bindings        =   "frmARTICO1.frx":7413
      DataField       =   "misura"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   11
      Top             =   3900
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Pattern"
      Text            =   "PZ"
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Bindings        =   "frmARTICO1.frx":7428
      DataField       =   "iva"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Top             =   5340
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Size"
      Text            =   ""
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   240
      Top             =   6000
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
   Begin MSDataListLib.DataCombo DataCombo4 
      Bindings        =   "frmARTICO1.frx":743D
      DataField       =   "iva1"
      DataSource      =   "datPrimaryRS"
      Height          =   315
      Left            =   2040
      TabIndex        =   16
      Top             =   5700
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Size"
      Text            =   ""
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1020
      Left            =   0
      ScaleHeight     =   1020
      ScaleWidth      =   15270
      TabIndex        =   31
      Top             =   7695
      Width           =   15270
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BarCode"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Code"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   52
      Top             =   320
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Produit"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   51
      Top             =   640
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Qtè en Stock"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   50
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix de Vente"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   49
      Top             =   1260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Dernière Sortie"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   48
      Top             =   2565
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Dernière Entrée"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   47
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fournisseur"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   46
      Top             =   3195
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock Maxi"
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   45
      Top             =   3540
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unité de mesure"
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   44
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Achat"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   43
      Top             =   4260
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stock Mini"
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   42
      Top             =   4980
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rechercher par CodeBarre"
      Height          =   255
      Left            =   3360
      TabIndex        =   41
      Top             =   5460
      Width           =   2055
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Unit Cost 2"
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   40
      Top             =   1620
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TVA"
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   39
      Top             =   5340
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Percentuale di Ricarico"
      Height          =   255
      Index           =   14
      Left            =   3240
      TabIndex        =   38
      Top             =   1620
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "% sur prix d'achat"
      Height          =   255
      Index           =   15
      Left            =   3240
      TabIndex        =   37
      Top             =   1260
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Carte Fidelitè"
      Height          =   255
      Index           =   16
      Left            =   2640
      TabIndex        =   36
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Service"
      Height          =   255
      Index           =   17
      Left            =   120
      TabIndex        =   35
      Top             =   2220
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Prix d'Achat Escompté"
      Height          =   255
      Index           =   18
      Left            =   120
      TabIndex        =   34
      Top             =   4620
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Category"
      Height          =   255
      Index           =   19
      Left            =   3240
      TabIndex        =   33
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tax 2"
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   32
      Top             =   5700
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "FRMARTICO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim stringa As String
Picture1.Picture = Nothing

Form6.Show vbModal
On Error Resume Next
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward


  stringa = txtFields(1) & ".jpg"
 Picture1.Picture = LoadPicture(stringa)
End Sub

Private Sub Command2_Click()
Text1.text = ""
Text1.SetFocus

End Sub

Private Sub Command3_Click()
frmricerca.Show
        datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceinterno = '" & art1 & "'", 0, adSearchForward

End Sub

Private Sub Command4_Click()
On Error GoTo dopo
Dim stringa As String
stringa = txtFields(1) & ".jpg"

SavePicture Picture1.Picture, stringa

dopo:
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
'rtb.Text = ""
Picture1.Cls

File1.Pattern = "*.jpg;*.bmp"

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

Private Sub Form_Load()
Command1_Click
End Sub
Private Sub Form_Activate()
File1.Pattern = "*.jpg;*.bmp"
lingue
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub datPrimaryRS_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
  'Posizione in cui inserire il codice per la gestione degli errori
  'Per ignorare gli errori, impostare come commento la riga seguente
  'Per intercettare gli errori, inserire il codice per la gestione degli errori in questa posizione
  MsgBox "Data error event hit err:" & Description
End Sub

Private Sub datPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Visualizza la posizione del record corrente per questo gruppo di record
  datPrimaryRS.Caption = "Record: " & CStr(datPrimaryRS.Recordset.AbsolutePosition)
End Sub

Private Sub datPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'Posizione in cui inserire il codice per la convalida
  'L'evento viene richiamato in seguito alle seguenti azioni
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
  datPrimaryRS.Recordset.AddNew
  Picture1 = Nothing
  
  Exit Sub
AddErr:
  MsgBox Err.Description
End Sub

Private Sub cmdDelete_Click()
Dim stringa As String
response = MsgBox(messaggi(9), vbOKCancel + vbCancel, messaggi(8))
Select Case response
 Case 6
On Error GoTo DeleteErr
  With datPrimaryRS.Recordset
    .delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  stringa = txtFields(1) & ".jpg"
Kill stringa
  


 Case Else
 
 End Select
Exit Sub
DeleteErr:
  MsgBox Err.Description
End Sub

Private Sub cmdRefresh_Click()
  'Necessario solo per applicazioni multiutente
  On Error GoTo RefreshErr
  datPrimaryRS.Refresh
  Exit Sub
RefreshErr:
  MsgBox Err.Description
End Sub

Private Sub cmdUpdate_Click()
  On Error GoTo UpdateErr

  datPrimaryRS.Recordset.UpdateBatch adAffectAll
  Exit Sub
UpdateErr:
  MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Label2_Click()

End Sub

Private Sub Text1_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
               datPrimaryRS.Recordset.MoveFirst
  datPrimaryRS.Recordset.Find "codiceean = '" & Text1.text & "'", 0, adSearchForward

     End Select

End Sub


Private Sub txtFields_Change(Index As Integer)

Dim numero1, numero2, numero3, numero4 As Single

On Error Resume Next
 
     
  numero1 = (txtFields(10))
  numero2 = (txtFields(12))
  'numero3 = (txtFields(13))
  numero4 = (txtFields(9))
  
  
 Select Case Index
    Case 1
     
  stringa = txtFields(1) & ".jpg"
 Picture1.Picture = LoadPicture(stringa)
  
        Case 10, 9, 12
        
         If txtFields(12) > 0 Then
         txtFields(7) = Format(((numero1 * (numero2) / 100)) + (numero1), "####0.00")
             End If
             
            
             
             ' If txtFields(9) > 0 And txtFields(9) <> "" Then
        ' txtFields(4) = Format(((numero1 * (numero4) / 100)) + (numero1), "####0.00")
         '    End If
             
     Case Else
     End Select
     
End Sub
Public Sub lingue()

Select Case lingua
  
 Case Is = "1F"
   Exit Sub
   
   Case Is = "2I"
lblLabels(0) = "Barcode"
lblLabels(1) = "Codice"
lblLabels(2) = "Prodotto"
lblLabels(3) = "Qtà in Stock"
lblLabels(4) = "Prezzo di Vendita"
lblLabels(15) = "% di ricarico"
lblLabels(17) = "Servizio"
lblLabels(16) = "Carta Fedeltà"
lblLabels(5) = "Data Ultima Vendita"
lblLabels(6) = "Data Ultimo Acquisto"
lblLabels(7) = "Fornitore"
lblLabels(8) = "Qtà Riordino"
lblLabels(9) = "Unità di Misura"
lblLabels(10) = "Costo"
lblLabels(18) = "Costo scontato"
lblLabels(11) = "Qtà Minima"
lblLabels(13) = "IVA"
Label1 = "Ricerca per Barcode"

Command2.Caption = "Nuova Ricerca"
cmdAdd.Caption = "Nuovo"
cmdUpdate.Caption = "Salva"
cmdDelete.Caption = "Elimina"
cmdRefresh.Caption = "Aggiorna"
cmdClose.Caption = "Esci"

Command4.Caption = " Memorizza Immagine"
FRMARTICO.Caption = "Archivo Articoli"


 Case Is = "3G"
 lblLabels(0) = "Barcode"
lblLabels(1) = "Code"
lblLabels(2) = "Product"
lblLabels(3) = "Qty in Stock"
lblLabels(4) = "Sell Price"
lblLabels(15) = "% on cost"
lblLabels(17) = "Service"
lblLabels(16) = "Fidelity Card"
lblLabels(5) = "Last Sell Date"
lblLabels(6) = "Last Entry Date"
lblLabels(7) = "Vendor"
lblLabels(8) = "Reorder Qty"
lblLabels(9) = "Misure Unit"
lblLabels(10) = "Cost"
lblLabels(18) = "Discount Cost"
lblLabels(11) = "Reorder Level"
lblLabels(13) = "Tax"
Label1 = "Barcode Search"

Command2.Caption = "New Search"
cmdAdd.Caption = "New"
cmdUpdate.Caption = "Save"
cmdDelete.Caption = "Delete"
cmdRefresh.Caption = "Update"
cmdClose.Caption = "Exit"

Command4.Caption = "Save Image"
FRMARTICO.Caption = "Products"

Case Is = "4S"
 lblLabels(0) = "Barcode"
lblLabels(1) = "Codigo"
lblLabels(2) = "Producto"
lblLabels(3) = "Cantidad en Stock"
lblLabels(4) = "Precio"
lblLabels(15) = "% de recarga"
lblLabels(17) = "Servicio"
lblLabels(16) = "Tarjeta Fidelidad"
lblLabels(5) = "Fecha de Ultima Venta"
lblLabels(6) = "Fecha de Ultimo Ingreso"
lblLabels(7) = "Proveedor"
lblLabels(8) = "Cantidad de Pedido"
lblLabels(9) = "Unidad"
lblLabels(10) = "Coste"
lblLabels(18) = "Coste Descontado"
lblLabels(11) = "Existencia Minima"
lblLabels(13) = "Iva"
Label1 = "Busqueda por Codigo Barreado"

Command2.Caption = "Nueva Busqueda"
cmdAdd.Caption = "Nuevo"
cmdUpdate.Caption = "Guarda"
cmdDelete.Caption = "Borra"
cmdRefresh.Caption = "Actualiza"
cmdClose.Caption = "Salir"

Command4.Caption = "Guarda Imagen"
FRMARTICO.Caption = "Productos"

 Case Is = "5D"
 lblLabels(0) = "Barcode"
lblLabels(1) = "Code"
lblLabels(2) = "Artikel"
lblLabels(3) = "Lagerbestand"
lblLabels(4) = "Verkaufspreis Brutto"
lblLabels(15) = "+% Einkaufspreis Netto"
lblLabels(17) = "Dienste"
lblLabels(16) = "TreueKarte"
lblLabels(5) = "Datum letzter Verkauf"
lblLabels(6) = "Datum letzter Kauf"
lblLabels(7) = "Lieferant"
lblLabels(8) = "Meldebestand"
lblLabels(9) = "U.M."
lblLabels(10) = "Einkaufspreis Netto"
lblLabels(18) = "Preisnachlaßkosten"
lblLabels(11) = "Bestellmenge"
lblLabels(13) = "Mwst"
Label1 = "Barcode Suchen"

Command2.Caption = "Neu Search"
cmdAdd.Caption = "Neu"
cmdUpdate.Caption = "Übernehmen"
cmdDelete.Caption = "Löschen"
cmdRefresh.Caption = "Update"
cmdClose.Caption = "Beenden"

Command4.Caption = "Bild Übernehmen"
FRMARTICO.Caption = "Artikel"

Case Else
End Select
 
End Sub

