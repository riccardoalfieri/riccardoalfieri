VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Produits"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form5"
   ScaleHeight     =   9060
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   9960
      TabIndex        =   6
      Top             =   7680
      Width           =   1935
      Begin VB.CheckBox Check2 
         Caption         =   "Check2"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Services"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Produits"
         Height          =   255
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Désignation Recherche"
      Height          =   855
      Left            =   8640
      MaskColor       =   &H00FF8080&
      Picture         =   "Form6.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Code Recherche"
      Height          =   855
      Left            =   1680
      MaskColor       =   &H00FF8080&
      Picture         =   "Form6.frx":1D12
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton command1 
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
      Height          =   855
      Left            =   11880
      Picture         =   "Form6.frx":3A24
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   7680
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3120
      TabIndex        =   0
      Top             =   7680
      Width           =   5775
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   2640
      Top             =   7200
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
      RecordSource    =   "select * from tires ORDER BY CodiceInterno,descrizione"
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Form6.frx":3E66
      Height          =   7695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   13573
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
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
      Caption         =   "Libellé des Produits"
      ColumnCount     =   21
      BeginProperty Column00 
         DataField       =   "descrizione"
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
         DataField       =   "Iva"
         Caption         =   "Tax"
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
         DataField       =   "listino1"
         Caption         =   "Price 1"
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
         DataField       =   "listino2"
         Caption         =   "Price 2"
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
         DataField       =   "listino3"
         Caption         =   "Price 3"
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
         DataField       =   "fornitore"
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
      BeginProperty Column06 
         DataField       =   "misura"
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
      BeginProperty Column07 
         DataField       =   "codiceEAN"
         Caption         =   "Barcode"
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
         DataField       =   "CodiceInterno"
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
      BeginProperty Column10 
         DataField       =   "Pattern"
         Caption         =   "Pattern"
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
         DataField       =   "Size"
         Caption         =   "Size"
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
         DataField       =   "reparto"
         Caption         =   "giorni"
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
         DataField       =   "Varieta"
         Caption         =   "Varieta"
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
         DataField       =   "Giacenza"
         Caption         =   "Giacenza"
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
         DataField       =   "ScortaMinima"
         Caption         =   "ScortaMinima"
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
         DataField       =   "LottoRiordino"
         Caption         =   "LottoRiordino"
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
         DataField       =   "pallet"
         Caption         =   "pallet"
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
      BeginProperty Column18 
         DataField       =   "PrezzoAcquisto"
         Caption         =   "PrezzoAcquisto"
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
         DataField       =   "DataUltimoCarico"
         Caption         =   "DataUltimoCarico"
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
      BeginProperty Column20 
         DataField       =   "DataUltimaVendita"
         Caption         =   "DataUltimaVendita"
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
      SplitCount      =   2
      BeginProperty Split0 
         Size            =   90
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column05 
            Alignment       =   2
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column15 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column16 
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column17 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column19 
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column20 
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
      BeginProperty Split1 
         Size            =   458
         BeginProperty Column00 
            Object.Visible         =   -1  'True
            ColumnWidth     =   4004,788
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   -1  'True
            ColumnWidth     =   3000,189
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   750,047
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275,024
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1365,165
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column18 
            Alignment       =   1
            Object.Visible         =   -1  'True
            ColumnWidth     =   1200,189
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
            ColumnWidth     =   2429,858
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command3_Click()
On Error Resume Next
Dim stringa As String

Select Case Check1.Value
Case 1
  If Check2.Value = 0 Then
 stringa = "SELECT * FROM tires WHERE [codiceinterno] like '" & Text1 & "%' and service=false"
     Else
      stringa = "SELECT * FROM tires WHERE [codiceinterno] like '" & Text1 & "%'"
    End If
    
    Case 0
     If Check2.Value = 0 Then
       stringa = "SELECT * FROM tires WHERE [codiceinterno] like '" & Text1 & "%'"
       Else
        stringa = "SELECT * FROM tires WHERE [codiceinterno] like '" & Text1 & "%' and service=true"
       End If
       Case Else
       End Select
       
    
   With Adodc2
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid2
     .ClearFields
     .HoldFields
     .ReBind
     End With
End Sub

Private Sub Command4_Click()
On Error Resume Next
Dim stringa As String

Select Case Check1.Value
Case 1
  If Check2.Value = 0 Then
 stringa = "SELECT * FROM tires WHERE [descrizione] like '" & Text2 & "%' and service=false"
     Else
      stringa = "SELECT * FROM tires WHERE [descrizione] like '" & Text2 & "%'"
    End If
    
    Case 0
     If Check2.Value = 0 Then
       stringa = "SELECT * FROM tires WHERE [descrizione] like '" & Text2 & "%'"
       Else
        stringa = "SELECT * FROM tires WHERE [descrizione] like '" & Text2 & "%' and service=true"
       End If
       Case Else
       End Select
       
    
   With Adodc2
    .RecordSource = stringa
    .Refresh
    End With
    
    With DataGrid2
     .ClearFields
     .HoldFields
     .ReBind
     End With
End Sub

Private Sub DataGrid2_Click()
DataGrid2_DblClick
End Sub

Private Sub DataGrid2_DblClick()
'On Error Resume Next
art1 = DataGrid2.Columns(8)
art2 = DataGrid2.Columns(0)
art3 = DataGrid2.Columns(13)
art4 = DataGrid2.Columns(2)
  
   art9 = DataGrid2.Columns(3)
       
   
 art10 = DataGrid2.Columns(4)
   
       
      
art5 = DataGrid2.Columns(6)
art6 = DataGrid2.Columns(14)
art7 = DataGrid2.Columns(18)
art8 = DataGrid2.Columns(1)
Unload Me
End Sub

Private Sub Form_Load()
lingue
Check1.Value = 1
End Sub

Private Sub Text1_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
        Command3_Click

     End Select

End Sub

Private Sub Text2_KeyDown( _
           KeyCode As Integer, Shift As Integer)

     Select Case KeyCode
     Case vbKeyReturn:
        Command4_Click

     End Select

End Sub
Private Sub lingue()

Select Case lingua
  
 Case Is = "1F"
   Exit Sub
   
   Case Is = "2I"
Form6.Caption = "Prodotti"
Command3.Caption = "Ricerca per codice"
Command4.Caption = "Ricerca per descrizione"
Command1.Caption = "Esci"
DataGrid2.Caption = "Prodotti"
DataGrid2.Columns(8).Caption = "Codice"
DataGrid2.Columns(0).Caption = "Prodotto"
DataGrid2.Columns(5).Caption = "Fornitore"
DataGrid2.Columns(18).Caption = "Costo"
Label1.Caption = "Prodotti"
Label2.Caption = "Servizi"
Frame1.Caption = "Scelta:"

 Case Is = "3G"
Form6.Caption = "Products"
Command4.Caption = "Code Search"
Command3.Caption = "Description Search"
Command1.Caption = "Exit"
DataGrid2.Caption = "Products"
DataGrid2.Columns(8).Caption = "Code"
DataGrid2.Columns(0).Caption = "Product"
Label1.Caption = "Products"
Label2.Caption = "Services"
Frame1.Caption = "Choose:"

Case Is = "4S"
Form6.Caption = "Lista de los Productos"
Command4.Caption = "Buscar por Codigo"
Command3.Caption = "Buscar por Descripcion"
Command1.Caption = "Salir"
DataGrid2.Caption = "Productos"
DataGrid2.Columns(8).Caption = "Codigo"
DataGrid2.Columns(0).Caption = "Producto"
Label1.Caption = "Productos"
Label2.Caption = "Servicios"
Frame1.Caption = "Elección:"

 Case Is = "5D"
Form6.Caption = "Artikel"
Command4.Caption = "Code Suchen"
Command3.Caption = "Bez Suchen"
Command1.Caption = "Beenden"
DataGrid2.Caption = "Artikel"
DataGrid2.Columns(8).Caption = "Code"
DataGrid2.Columns(0).Caption = "Artikel"
Label1.Caption = "Artikel"
Label2.Caption = "Dienste"
Frame1.Caption = "Wählen Sie:"

 Case Else
 End Select
End Sub

