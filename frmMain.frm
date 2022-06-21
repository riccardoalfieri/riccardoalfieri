VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "CAISSE 3-4-5"
   ClientHeight    =   7125
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13185
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMain.frx":0442
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   1535
      ButtonWidth     =   2223
      ButtonHeight    =   1429
      AllowCustomize  =   0   'False
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   9
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Produits"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Agenda"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Clients"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Collaborateurs"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "File d'Attente"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   26
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Caisse"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   27
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Encaissements"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   28
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "Accès"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   23
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Quitter"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   1680
      Top             =   3240
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1920
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   4440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   6750
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "www.axasoft.org"
            TextSave        =   "www.axasoft.org"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2646
            MinWidth        =   2646
            Text            =   "Multilanguage"
            TextSave        =   "Multilanguage"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "axasoft@altervista.org"
            TextSave        =   "axasoft@altervista.org"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   28
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":292E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3580
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":41D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":4E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":5A76
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":66C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":731A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":7F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":8BBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":9810
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":A462
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":B0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":BD06
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":C958
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":D5AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":E1FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":EE4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":FAA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":106F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11344
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":11F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":12BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1383A
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1448C
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":150DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":15D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":16982
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&Société"
      Begin VB.Menu società 
         Caption         =   "&Société"
      End
      Begin VB.Menu mnusetting 
         Caption         =   ""
      End
      Begin VB.Menu mnuticket 
         Caption         =   ""
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnuarchivi 
      Caption         =   "&Fichiers"
      WindowList      =   -1  'True
      Begin VB.Menu mnuArchiviArticoli 
         Caption         =   "&Produits"
      End
      Begin VB.Menu mnuArchivioClienti 
         Caption         =   "Clients"
      End
      Begin VB.Menu mnuArchivioFornitori 
         Caption         =   "&Fornisseurs"
      End
      Begin VB.Menu mnuArchivioIva 
         Caption         =   "&Codes TVA"
      End
      Begin VB.Menu mnuarchiviofarmaci 
         Caption         =   ""
      End
      Begin VB.Menu mnuarchiviopatologie 
         Caption         =   "Formula"
      End
      Begin VB.Menu mnuarchivioposologie 
         Caption         =   "&Collaborateur"
      End
      Begin VB.Menu mnucaisse 
         Caption         =   "Cai&sse"
      End
   End
   Begin VB.Menu mnuapp 
      Caption         =   "Ventes"
      Begin VB.Menu mnuappapp 
         Caption         =   "&Agenda"
      End
      Begin VB.Menu mnuappqueue 
         Caption         =   "&File d'Attente"
      End
      Begin VB.Menu mnuapppos 
         Caption         =   "&Sortie"
      End
      Begin VB.Menu mnuventessoldes 
         Caption         =   "Soldes &clientes"
      End
   End
   Begin VB.Menu mnuAcquisti 
      Caption         =   "Achats"
      Begin VB.Menu mnuOrdineFornitore 
         Caption         =   "&Commandes d'Achats"
      End
      Begin VB.Menu mnuAggiornaFornitore 
         Caption         =   "&Ajourner le Stock"
      End
      Begin VB.Menu mnuachatssoldes 
         Caption         =   "Soldes &Fournisseur"
      End
      Begin VB.Menu mnuachatsecheancier 
         Caption         =   ""
      End
   End
   Begin VB.Menu mnumagazzino 
      Caption         =   "&Editions"
      Begin VB.Menu mnumagazzinocarico 
         Caption         =   "&Ma Caisse"
      End
      Begin VB.Menu mnusalesclient 
         Caption         =   "Encaissement"
      End
      Begin VB.Menu mnumagazzinoinventario 
         Caption         =   "&Inventaire "
      End
      Begin VB.Menu mnumagazzinofornitore 
         Caption         =   "Inventaire Fournisseur"
      End
      Begin VB.Menu mnureportreorder 
         Caption         =   "Produits en Rupture de Stock"
      End
      Begin VB.Menu mnuporders 
         Caption         =   "Liste des Commandes d'Achats"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&Point de Vente 2009"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Const EM_UNDO = &HC7
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Private Declare Function GetVolumeInformation Lib _
"kernel32.dll" Alias "GetVolumeInformationA" (ByVal _
lpRootPathName As String, ByVal lpVolumeNameBuffer As _
String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength _
As Long, lpFileSystemFlags As Long, ByVal _
lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long

Function GetSerialNumber(strDrive As String) As Long
Dim SerialNum As Long
Dim Res As Long
Dim Temp1 As String
Dim Temp2 As String
Temp1 = String$(255, Chr$(0))
Temp2 = String$(255, Chr$(0))
Res = GetVolumeInformation(strDrive, Temp1, _
Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
GetSerialNumber = SerialNum
End Function


Private Sub Command1_Click()

End Sub

Private Sub frmtires_Click()

End Sub

Private Sub frmstampadifferite_Click()
frmstampadiff.Show
End Sub

Private Sub MDIForm_Load()
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
 
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=db1.mdb;Persist Security Info=False;"
   cn.Open
   
     Dim a, filepath As String
     
  codice_avvio = "ZXCVBNM,.-"
   serialesoftware = "PointVente" & GetSerialNumber("C:\")
     Timer1.Enabled = True
      Timer1.Interval = 25000
      
      
     LoadNewDoc


lingue
   
 
End Sub


Private Sub LoadNewDoc()


    Static lDocumentCount As Long
    Dim frmD As frmDocument
    lDocumentCount = lDocumentCount + 1
    Set frmD = New frmDocument
  frmD.Caption = "Document " & lDocumentCount
     frmD.Picture = LoadPicture("desktop.jpg")
       
    frmD.Show
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    cn.Close
    End
End Sub

Private Sub mnuachatssoldes_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmpartitari.Show
End Sub

Private Sub mnuAggiornaFornitore_Click()
Shell "AggiornaStock.exe", vbNormalFocus
End Sub

Private Sub mnuappapp_Click()
Shell "miniagenda.exe"
End Sub

Private Sub mnuapppos_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If


Shell "pos.exe " & codice_avvio, vbNormalFocus
End Sub

Private Sub mnuappqueue_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

End Sub



Private Sub mnuArchiviArticoli_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
Shell "articoli.exe " & codice_avvio, vbNormalFocus
End Sub

Private Sub mnuArchivioClienti_Click()
Shell "client.exe " & codice_avvio, vbNormalFocus

End Sub

Private Sub mnuArchivioFornitori_Click()
Shell "fornitori.exe", vbNormalFocus
End Sub

Private Sub mnuArchivioIva_Click()
frmSize.Show
End Sub

Private Sub mnuArchivioMisure_Click()
frmPattern.Show
End Sub

Private Sub mnuarchiviopatologie_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmlistaposologia.Show
End Sub
Private Sub mnuarchivioposologie_Click()
frmcollaboratori.Show
End Sub
Private Sub mnuArchivioPagamenti_Click()
frmArchivioPagamenti.Show

End Sub

Private Sub mnuDdt_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmddtfornitore.Show
End Sub

Private Sub mnuFattureFornitori_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmfatturefornitore.Show
End Sub



Private Sub mnuinventarioreale_Click()
Set DataEnvironment5 = New DataEnvironment5
DataReport5.Show
End Sub

Private Sub mnucaisse_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

frmcassa.Show
End Sub

Private Sub mnufidel_Click()
Set DataEnvironment8 = New DataEnvironment8

With DataReport8
     Select Case lingua
 Case Is = "1F"
 
   
   Case Is = "2I"
     With .Sections("Section4").Controls
                      .Item("label2").Caption = "Carta Fedeltà"
     End With
     With .Sections("Section2").Controls
              .Item("label7").Caption = "Codice"
              .Item("label1").Caption = "Cliente"
                         .Item("label3").Caption = "Telefono"
                          .Item("label5").Caption = "Totale"
     End With
          Case Is = "3G"
     With .Sections("Section4").Controls
                           .Item("label2").Caption = "Fidelity Card List"
     End With
     With .Sections("Section2").Controls
                   .Item("label7").Caption = "Code"
              .Item("label1").Caption = "Client"
                         .Item("label3").Caption = "Phone"
                          .Item("label5").Caption = "Total"
     End With
        Case Is = "4S"
     With .Sections("Section4").Controls
                           .Item("label2").Caption = "Tarjeta Fidelidad"
     End With
     With .Sections("Section2").Controls
                   .Item("label7").Caption = "Codigo"
              .Item("label1").Caption = "Cliente"
                         .Item("label3").Caption = "Telefono"
                          .Item("label5").Caption = "Total"
     End With
     
       Case Is = "5D"
     With .Sections("Section4").Controls
                           .Item("label2").Caption = "Treuekarte"
     End With
     With .Sections("Section2").Controls
                   .Item("label7").Caption = "Code"
              .Item("label1").Caption = "Kunde"
                         .Item("label3").Caption = "Telefon"
                          .Item("label5").Caption = "Gesamt"
     End With
     Case Else
     End Select
     
 
    .Show
 End With
End Sub

Private Sub mnumagazzinocarico_Click()
Set DataEnvironment5 = New DataEnvironment5
DataEnvironment5.Command1 Date
With DataReport5
Select Case lingua
 Case Is = "1F"
 With .Sections("Section4").Controls
              .Item("label10").Caption = " " & Date
                  End With

   
   Case Is = "2I"
     With .Sections("Section4").Controls
              .Item("label10").Caption = " " & Date
              .Item("label2").Caption = "Incassi"
     End With
     With .Sections("Section2").Controls
              .Item("label7").Caption = "Data"
             ' .Item("label1").Caption = "Cliente"
              .Item("label8").Caption = "Pagamento"
             ' .Item("label3").Caption = "Subtotale"
             ' .Item("label4").Caption = "Sconto"
              .Item("label5").Caption = "Totale"
     End With
       With .Sections("Section5").Controls
              .Item("label6").Caption = "Totale"
     End With
    Case Is = "3G"
     With .Sections("Section4").Controls
              .Item("label10").Caption = " " & Date
              .Item("label2").Caption = "Sales"
     End With
     With .Sections("Section2").Controls
              .Item("label7").Caption = "Date"
            '  .Item("label1").Caption = "Client"
              .Item("label8").Caption = "Payment"
            '  .Item("label3").Caption = "Subtotal"
            ' .Item("label4").Caption = "Discount"
              .Item("label5").Caption = "Total"
     End With
       With .Sections("Section5").Controls
              .Item("label6").Caption = "Total"
     End With
     Case Is = "4S"
     With .Sections("Section4").Controls
              .Item("label10").Caption = " " & Date
              .Item("label2").Caption = "CAJA"
     End With
     With .Sections("Section2").Controls
              .Item("label7").Caption = "Fecha"
              '.Item("label1").Caption = "Cliente"
              .Item("label8").Caption = "Pago"
             ' .Item("label3").Caption = "Subtotal"
            '  .Item("label4").Caption = "Desc.%"
              .Item("label5").Caption = "Total"
     End With
       With .Sections("Section5").Controls
              .Item("label6").Caption = "Total"
     End With
     Case Is = "5D"
     With .Sections("Section4").Controls
              .Item("label10").Caption = " " & Date
              .Item("label2").Caption = "Kasse"
     End With
     With .Sections("Section2").Controls
              .Item("label7").Caption = "Datum"
             ' .Item("label1").Caption = "Kunde"
              .Item("label8").Caption = "Zahlung"
            '  .Item("label3").Caption = "Netto"
            '  .Item("label4").Caption = "Rabatt"
              .Item("label5").Caption = "Brutto"
     End With
       With .Sections("Section5").Controls
              .Item("label6").Caption = "Gesamte"
     End With
     Case Else
     End Select
     
    .Show
 End With
End Sub

Private Sub mnumagazzinofornitore_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
Shell "InventarioFornitori.exe", vbNormalFocus
End Sub

Private Sub mnumagazzinoinventario_Click()
Shell "Inventario.exe", vbNormalFocus
End Sub


Private Sub mnuOrdineFornitore_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
Shell "OrdiniFornitori.exe", vbNormalFocus
End Sub

Private Sub mnuOrdiniModifica_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmAggiorna.Show
End Sub



Private Sub mnuricevute_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

frmricevute.Show
End Sub

Private Sub mnuristampe_Click()
frmristampe.Show
End Sub

Private Sub mnuporders_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
Shell "StampaAcquisti.exe", vbNormalFocus
End Sub

Private Sub mnupurchase_Click()
frmstampa.Show
End Sub

Private Sub mnusalesclient_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmstampevarie1.Show
End Sub

Private Sub mnureportreorder_Click()

Shell "riordino.exe", vbNormalFocus

End Sub

Private Sub mnuSchedeFornitori_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
mnufidel_Click
End Sub

Private Sub mnuschedescadenziario_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

frmscadenzario.Show
End Sub

Private Sub mnustampaddt_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmstampaddt.Show
End Sub

Private Sub mnuStampeFatture_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
frmpos.Show
End Sub
Private Sub mnuStampeOrdini_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If
  frmstampa.Show

End Sub


Private Sub mnusetting_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

End Sub

Private Sub mnuticket_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

End Sub

Private Sub mnuventessoldes_Click()
If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If

frmcontafat.Show
End Sub

Private Sub società_Click()
frmAzienda.Show
End Sub



Private Sub mnuHelpAbout_Click()
    MsgBox messaggi(11)
End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'Se il progetto non include un file della Guida, visualizza un messaggio per
    'l'utente. È possibile impostare il file della Guida per l'applicazione nella
    'finestra di dialogo Proprietà progetto.
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossibile visualizzare il Sommario della Guida. Nessun file della Guida associato al progetto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'Se il progetto non include un file della Guida, visualizza un messaggio per
    'l'utente. È possibile impostare il file della Guida per l'applicazione nella
    'finestra di dialogo Proprietà progetto.
    If Len(App.HelpFile) = 0 Then
        MsgBox "Impossibile visualizzare il Sommario della Guida. Nessun file della Guida associato al progetto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub



Private Sub mnuViewWebBrowser_Click()
    'Da fare: Aggiunge il codice per 'mnuViewWebBrowser_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewWebBrowser_Click'."
End Sub

Private Sub mnuViewOptions_Click()
    'Da fare: Aggiunge il codice per 'mnuViewOptions_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewOptions_Click'."
End Sub

Private Sub mnuViewRefresh_Click()
    'Da fare: Aggiunge il codice per 'mnuViewRefresh_Click'.
    MsgBox "Aggiunge il codice per 'mnuViewRefresh_Click'."
End Sub

Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sbStatusBar.Visible = mnuViewStatusBar.Checked
End Sub

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Toolbar1.Visible = mnuViewToolbar.Checked
End Sub

Private Sub mnuEditPasteSpecial_Click()
    'Da fare: Aggiunge il codice per 'mnuEditPasteSpecial_Click'.
    MsgBox "Aggiunge il codice per 'mnuEditPasteSpecial_Click'."
End Sub

Private Sub mnuEditPaste_Click()
    On Error Resume Next
    ActiveForm.rtfText.SelRTF = Clipboard.GetText

End Sub

Private Sub mnuEditCopy_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF

End Sub

Private Sub mnuEditCut_Click()
    On Error Resume Next
    Clipboard.SetText ActiveForm.rtfText.SelRTF
    ActiveForm.rtfText.SelText = vbNullString

End Sub

Private Sub mnuEditUndo_Click()
    'Da fare: Aggiunge il codice per 'mnuEditUndo_Click'.
    MsgBox "Aggiunge il codice per 'mnuEditUndo_Click'."
End Sub


Private Sub mnuFileExit_Click()
    'Scarica il form.
    Unload Me

End Sub

Private Sub mnuFileSend_Click()
    'Da fare: Aggiunge il codice per 'mnuFileSend_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileSend_Click'."
End Sub

Private Sub mnuFilePrint_Click()
    On Error Resume Next
    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Stampa"
        .CancelError = True
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If ActiveForm.rtfText.SelLength = 0 Then
            .Flags = .Flags + cdlPDAllPages
        Else
            .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        If Err <> MSComDlg.cdlCancel Then
            ActiveForm.rtfText.SelPrint .hDC
        End If
    End With

End Sub

Private Sub mnuFilePrintPreview_Click()
    'Da fare: Aggiunge il codice per 'mnuFilePrintPreview_Click'.
    MsgBox "Aggiunge il codice per 'mnuFilePrintPreview_Click'."
End Sub

Private Sub mnuFilePageSetup_Click()
    On Error Resume Next
    With dlgCommonDialog
        .DialogTitle = "Imposta pagina"
        .CancelError = True
        .ShowPrinter
    End With

End Sub

Private Sub mnuFileProperties_Click()
    'Da fare: Aggiunge il codice per 'mnuFileProperties_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileProperties_Click'."
End Sub

Private Sub mnuFileSaveAll_Click()
    'Da fare: Aggiunge il codice per 'mnuFileSaveAll_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileSaveAll_Click'."
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim sFile As String
    

    If ActiveForm Is Nothing Then Exit Sub
    

    With dlgCommonDialog
        .DialogTitle = "Salva con nome"
        .CancelError = False
        'Da fare: impostare i flag e gli attributi del controllo CommonDialog.
        .Filter = "Tutti i file (*.*)|*.*"
        .ShowSave
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    ActiveForm.Caption = sFile
    ActiveForm.rtfText.SaveFile sFile

End Sub

Private Sub mnuFileSave_Click()
    Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With dlgCommonDialog
            .DialogTitle = "Salva"
            .CancelError = False
            'Da fare: impostare i flag e gli attributi del controllo CommonDialog.
            .Filter = "Tutti i file (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
    Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub

Private Sub mnuFileClose_Click()
    'Da fare: Aggiunge il codice per 'mnuFileClose_Click'.
    MsgBox "Aggiunge il codice per 'mnuFileClose_Click'."
End Sub


Private Sub mnuFileNew_Click()
    LoadNewDoc
End Sub


Public Sub lingue()

Select Case lingua
  
 Case Is = "1F"
 
  messaggi(1) = "Supprimer?"
messaggi(2) = "Quitter?"
messaggi(3) = "Paiement!!!"
messaggi(4) = "Valider?"
messaggi(5) = " Annuler?"
messaggi(6) = "Inserer le client"
 messaggi(7) = " Insérer un numéro "
messaggi(8) = "Attention"
 messaggi(9) = "Effacer la ligne?"
 messaggi(10) = " Impostare la Commport"
 messaggi(11) = "CAISSE 3-4-5 - www.axasoft.org axasoft@altervista.org Ver. " & App.Major & "." & App.Minor & "." & App.Revision
 messaggi(12) = "Total il est 0"
 
Me.Caption = "CAISSE 3-4-5"
Toolbar1.Buttons(1).Caption = "Produits"
Toolbar1.Buttons(2).Caption = "Agenda"

Toolbar1.Buttons(3).Caption = "Clients"
Toolbar1.Buttons(4).Caption = "Collaborateurs"
Toolbar1.Buttons(5).Caption = "Caisse Set"
Toolbar1.Buttons(6).Caption = "Caisse"
Toolbar1.Buttons(7).Caption = "Encaissements "
Toolbar1.Buttons(8).Caption = "Categories"
Toolbar1.Buttons(9).Caption = "Quitter"

mnuHelpAbout.Caption = "CAISSE 3-4-5"

 
   Exit Sub
   
   Case Is = "2I"
   
Me.Caption = "Punto Vendita 2009"
Toolbar1.Buttons(1).Caption = "Prodotti"
Toolbar1.Buttons(2).Caption = "Agenda"

Toolbar1.Buttons(3).Caption = "Clienti"
Toolbar1.Buttons(4).Caption = "Collaboratori"
Toolbar1.Buttons(5).Caption = "Settaggio"
Toolbar1.Buttons(6).Caption = "Cassa"
Toolbar1.Buttons(7).Caption = "Incassi "
Toolbar1.Buttons(8).Caption = "Categorie"
Toolbar1.Buttons(9).Caption = "Esci"

mnuFile.Caption = "File"
società.Caption = "Società"
mnuFileExit.Caption = "Esci"
mnuarchivi.Caption = "Archivi"
mnuArchiviArticoli.Caption = "Articoli"
mnuArchivioClienti.Caption = "Clienti"
mnuArchivioFornitori.Caption = "Fornitori"
'mnuArchivioMisure.Caption = "Unità di Misura"
mnuArchivioIva.Caption = "IVA"
'mnuarchiviofarmaci.Caption = "Condizioni di Pagamento"
mnuarchiviopatologie.Caption = "Formula"
mnuarchivioposologie.Caption = "Collaboratori"
mnucaisse.Caption = "Cassa"
mnuapp.Caption = "Vendite"
mnuappapp.Caption = "Agenda"
mnuappqueue.Caption = "Coda di lavoro"
mnuapppos.Caption = "Cassa"
mnuventessoldes.Caption = "Conto Clienti"
mnuAcquisti.Caption = "Acquisti"
mnuOrdineFornitore.Caption = "Ordini Fornitori"
mnuAggiornaFornitore.Caption = "Aggiorna Stock"
mnuachatssoldes.Caption = "Conto Fornitori"
'mnuachatsecheancier.Caption = "Scadenzario Fornitori"
mnumagazzino.Caption = "Stampe"
mnumagazzinocarico.Caption = "Cassa di Oggi"
mnusalesclient.Caption = "Incassi per periodo"
mnumagazzinoinventario.Caption = "Inventario"
mnumagazzinofornitore.Caption = "Inventario per Fornitore"
mnureportreorder.Caption = "Lista Riordino"
mnuporders.Caption = "Lista Ordini Acquisto"
'mnufidel.Caption = "Punti Fedeltà"
mnuHelp.Caption = "&?"
mnuHelpAbout.Caption = "Punto Vendita 2009"

messaggi(1) = "Elimino?"
messaggi(2) = "Manca il pagamento, uscire?"
messaggi(3) = "Pagamento!!!"
messaggi(4) = "Registro?"
messaggi(5) = " Annullo?"
messaggi(6) = "Inserire il cliente"
messaggi(7) = "Inserire un numero [0 - 10000]"
messaggi(8) = "Attenzione"
messaggi(9) = "Cancellare la riga?"
messaggi(10) = " Impostare la Commport"
messaggi(11) = "Punto Vendita 2009 - www.axasoft.org axasoft@altervista.org Ver. " & App.Major & "." & App.Minor & "." & App.Revision
 messaggi(12) = "Il totale è 0"




 Case Is = "3G"
   
Me.Caption = "CAISSE 3-4-5"
Toolbar1.Buttons(1).Caption = "Products"
Toolbar1.Buttons(2).Caption = "Agenda"

Toolbar1.Buttons(3).Caption = "Clients"
Toolbar1.Buttons(4).Caption = "Employees"
Toolbar1.Buttons(5).Caption = "Pos Set"
Toolbar1.Buttons(6).Caption = "POS"
Toolbar1.Buttons(7).Caption = "Sales"
Toolbar1.Buttons(8).Caption = "Categories"
Toolbar1.Buttons(9).Caption = "Exit"

mnuFile.Caption = "File"
società.Caption = "Society"
mnuFileExit.Caption = "Exit"
mnuarchivi.Caption = "Archives"
mnuArchiviArticoli.Caption = "Products"
mnuArchivioClienti.Caption = "Clients"
mnuArchivioFornitori.Caption = "Vendors"
'mnuArchivioMisure.Caption = "Unit"
mnuArchivioIva.Caption = "TAX"
'mnuarchiviofarmaci.Caption = "Payments"
mnuarchiviopatologie.Caption = "Formulas"
mnuarchivioposologie.Caption = "Employees"
mnucaisse.Caption = "Sales"
mnuapp.Caption = "Appointments"
mnuappapp.Caption = "Agenda"
mnuappqueue.Caption = "Queue"
mnuapppos.Caption = "POS"
mnuventessoldes.Caption = "Clients Accounts"
mnuAcquisti.Caption = "Purchases"
mnuOrdineFornitore.Caption = "Vendors Orders"
mnuAggiornaFornitore.Caption = "Receiving Inventory"
mnuachatssoldes.Caption = "Vendors Accounts"
'mnuachatsecheancier.Caption = "Payment Terms for Vendors Invoices"
mnumagazzino.Caption = "Reports"
mnumagazzinocarico.Caption = "Today's Sales"
mnusalesclient.Caption = "Sales"
mnumagazzinoinventario.Caption = "Inventory"
mnumagazzinofornitore.Caption = "Vendors Inventory"
mnureportreorder.Caption = "Reorder Report"
mnuporders.Caption = "Purchase Orders List"
'mnufidel.Caption = "Fidelity Card"
mnuHelp.Caption = "&?"
mnuHelpAbout.Caption = "CAISSE 3-4-5"

messaggi(1) = "Delete?"
messaggi(2) = "Payment not present. Exit?"
messaggi(3) = "Payment!!!"
messaggi(4) = "Save?"
messaggi(5) = "Cancel?"
messaggi(6) = "Insert client"
messaggi(7) = "Insert a number [0 - 10000]"
messaggi(8) = "Warning"
messaggi(9) = "Delete the line?"
messaggi(10) = " Impostare la Commport"
messaggi(11) = "CAISSE 3-4-5 - www.axasoft.org axasoft@altervista.org Ver. " & App.Major & "." & App.Minor & "." & App.Revision
messaggi(12) = "Total it is 0"
Case Is = "4S"
   
Me.Caption = "CAISSE 3-4-5"
Toolbar1.Buttons(1).Caption = "Productos"
Toolbar1.Buttons(2).Caption = "Agenda"

Toolbar1.Buttons(3).Caption = "Clientes"
Toolbar1.Buttons(4).Caption = "Colaboradores"
Toolbar1.Buttons(5).Caption = "Caja Set"
Toolbar1.Buttons(6).Caption = "Caja"
Toolbar1.Buttons(7).Caption = "Ventas"
Toolbar1.Buttons(8).Caption = "Categorias"
Toolbar1.Buttons(9).Caption = "Salir"

mnuFile.Caption = "File"
società.Caption = "Sociedad"
mnuFileExit.Caption = "Salir"
mnuarchivi.Caption = "Archivos"
mnuArchiviArticoli.Caption = "Productos"
mnuArchivioClienti.Caption = "Clientes"
mnuArchivioFornitori.Caption = "Proveedors"
'mnuArchivioMisure.Caption = "Unidad de medida"
mnuArchivioIva.Caption = "Iva"
'mnuarchiviofarmaci.Caption = "Pago"
mnuarchiviopatologie.Caption = "Formulas"
mnuarchivioposologie.Caption = "Colaboradores"
mnucaisse.Caption = "Ventas"
mnuapp.Caption = "Citas"
mnuappapp.Caption = "Agenda"
mnuappqueue.Caption = "Fila"
mnuapppos.Caption = "Caja"
mnuventessoldes.Caption = "Cuento Clientes"
mnuAcquisti.Caption = "Compras"
mnuOrdineFornitore.Caption = "Ordenas a Proveedors"
mnuAggiornaFornitore.Caption = "Actualiza Stock"
mnuachatssoldes.Caption = "Cuento Proveedors"
'mnuachatsecheancier.Caption = "Terminos de Pago"
mnumagazzino.Caption = "Imprimir"
mnumagazzinocarico.Caption = "Ventas de Hoy"
mnusalesclient.Caption = "Ventas"
mnumagazzinoinventario.Caption = "Stock"
mnumagazzinofornitore.Caption = "Stock para Proveedor"
mnureportreorder.Caption = "Lista de reordeno"
mnuporders.Caption = "Lista de las Compras"
'mnufidel.Caption = "Tarjeta Fidelidad"
mnuHelp.Caption = "&?"
mnuHelpAbout.Caption = "CAISSE 3-4-5"

messaggi(1) = "Elimino?"
messaggi(2) = "Manca il pagamento, uscire?"
messaggi(3) = "Pagamento!!!"
messaggi(4) = "Registro?"
messaggi(5) = " ¿borrar la línea?"
messaggi(6) = "Manca Cliente"
messaggi(7) = "Insertar un número [0 - 10000]"
messaggi(8) = "Atención"
messaggi(9) = "¿borrar la línea?"
messaggi(10) = " Impostare la Commport"
messaggi(12) = "Total es 0"
messaggi(11) = "CAISSE 3-4-5 - www.axasoft.org axasoft@altervista.org Ver. " & App.Major & "." & App.Minor & "." & App.Revision

Case Is = "5D"
   
Me.Caption = "CAISSE 3-4-5"
Toolbar1.Buttons(1).Caption = "Artikel"
Toolbar1.Buttons(2).Caption = "Termine"

Toolbar1.Buttons(3).Caption = "Kunden"
Toolbar1.Buttons(4).Caption = "Personal"
Toolbar1.Buttons(5).Caption = "Kasse Set"
Toolbar1.Buttons(6).Caption = "Kasse"
Toolbar1.Buttons(7).Caption = "Kassenjournal"
Toolbar1.Buttons(8).Caption = "Categories"
Toolbar1.Buttons(9).Caption = "Exit"

mnuFile.Caption = "File"
società.Caption = "Gesellschaft"
mnuFileExit.Caption = "Beenden"
mnuarchivi.Caption = "Artikel"
mnuArchiviArticoli.Caption = "Artikel"
mnuArchivioClienti.Caption = "Kunden"
mnuArchivioFornitori.Caption = "Lieferant"
'mnuArchivioMisure.Caption = "Einheit der Maßnahme"
mnuArchivioIva.Caption = "Mwst"
'mnuarchiviofarmaci.Caption = "Zuhlang"
mnuarchiviopatologie.Caption = "Formula"
mnuarchivioposologie.Caption = "Personal"
mnucaisse.Caption = "Kasse"
mnuapp.Caption = "Verkäufe"
mnuappapp.Caption = "Termine"
mnuappqueue.Caption = "Warteschlange"
mnuapppos.Caption = "Kasse"
mnuventessoldes.Caption = "KundenBuchhaltung"
mnuAcquisti.Caption = "Kauft"
mnuOrdineFornitore.Caption = "Kaufauftrag"
mnuAggiornaFornitore.Caption = "Aktualiserung des Bestandes"
mnuachatssoldes.Caption = "LieferantBuchhaltung"
'mnuachatsecheancier.Caption = "Zahlung"
mnumagazzino.Caption = "Drucke"
mnumagazzinocarico.Caption = "Kasse von heute"
mnusalesclient.Caption = "Kasse"
mnumagazzinoinventario.Caption = "Inventurliste"
mnumagazzinofornitore.Caption = "LiferantBestandes"
mnureportreorder.Caption = "Lagerverwaltung Inventurliste"
mnuporders.Caption = "Kaufaugtraliste"
'mnufidel.Caption = "Treuekarte"
mnuHelp.Caption = "&?"
mnuHelpAbout.Caption = "CAISSE 3-4-5"
messaggi(1) = "Delete?"
messaggi(2) = "Payment not present. Exit?"
messaggi(3) = "Payment!!!"
messaggi(4) = "Save?"
messaggi(5) = "Cancel?"
messaggi(6) = "Insert client"
messaggi(7) = "Legen Sie eine Zahl ein[0 - 10000]"
messaggi(8) = "Warnung"
messaggi(9) = "Löschen Sie die Linie?"
messaggi(10) = "Set Commport, Please "
messaggi(12) = "Gesamtsumme sind 0?"
messaggi(11) = "CAISSE 3-4-5- www.axasoft.org axasoft@altervista.org Ver. " & App.Major & "." & App.Minor & "." & App.Revision

  Case Else
  End Select
  


End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)

If allerta_conn = True And isRegistered = False Then
Shell "topone.exe " & altercode, vbNormalFocus
Exit Sub
End If


Select Case Button.Index
Case 1:
Shell "articoli.exe " & codice_avvio, vbNormalFocus

Case 2:
Shell "miniagenda.exe", vbNormalFocus
Case 3:
Shell "client.exe " & codice_avvio, vbNormalFocus

Case 4:
frmcollaboratori.Show
Case 5:


Case 6:
Shell "pos.exe " & codice_avvio, vbNormalFocus

Case 7:
mnumagazzinocarico_Click
Case 8:
 
Case 9:

Unload Me
End Select
End Sub
Private Sub Timer1_Timer()

If isRegistered = False Then  'IL LANCIO DEL TOPONE PARTE SOLO SE NON E' REGISTRATO
Check2
End If

End Sub
Private Sub Check2()
Dim codicesoftware, updateString As String
Dim rilevazione As Integer

'MsgBox isRegistered

updateString = Inet1.OpenURL("http://axasoft.altervista.org/banner/software/users/" & serialesoftware & ".txt")
rilevazione = Val(updateString) 'VALORE DELLA STRINGA SCARICATA

Select Case Len(updateString)
'RILEVIAMO DELLE INFORMAZIONI DALLA LUNGHEZZA DELLA STRINGA SE NON RIUSCIAMO AD OTTENERE UN VALORE NUMERICO
   ' NEL CASO DI VALORE NUMERICO CONFRONTIAMO TRA IL VALORE RILEVATO E IL VALORE DEL CONTATORE

      Case Is = 0 ' CONNESSIONE ASSENTE: allerta_conn=true
      '            '
      
           MsgBox msg(2), vbInformation, msg(16)
              allerta_conn = True
           
            
            
        Case Is > 100  'NON ESISTE LA PAGINA ONLINE in questo caso si riceve una pagina vuota
                        ' IL PROGRAMMA SI FERMA
                        
         MsgBox msg(17), vbInformation, msg(16)
              allerta_conn = True
           End
           
           Case Is > 10 ' IN QUESTO CASO SI VALUTA UN MESSAGGIO DELL'EDITORE CHE SIA INFERIORE O UGUALE A 100 CARATTERI
                        ' SE FOSSE SUPERIORE A 100 VERREBBE VISUALIZZATO PRIMA
                        
               MsgBox updateString, vbOKOnly, msg(18)
             
             Case Else
                   If contatore = 0 Then
              'MsgBox contatore & "-" & rilevazione & "Updatestring:" & updatestring
              'CONTATORE VIENE ATTIVATO SE =0 E PRENDE IL VALORE DEL NUMERO LETTO
              
                   contatore = rilevazione 'QUESTO E' IL CASO DEL CICLO INIZIALE
                                           ' E QUINDI ESCE DALLA SUB
                    Exit Sub
                    
                        Else
                        
                              Select Case rilevazione
                                  Case Is > contatore ' IN QUESTO CASO TUTTO PROCEDE E SI VA AVANTI
                                                      ' ALLERTA_CONN SI SPEGNE
                                                      ' L'INTERVALLO SI ASSESTA
                                                      ' IL CONTATORE SI AGGIORNA
                                ' MsgBox contatore & "-" & rilevazione & "Updatestring:" & updatestring
                                   contatore = rilevazione
                                   allerta_conn = False
                                   Timer1.Interval = 35000
                                   
                                   
          
                        Case Is = contatore 'QUESTO E' IL CASO IN CUI LA CONNESSIONE SI E' FERMATA
                                            ' ALLERTA_CONN SI ATTIVA
                                            ' VIENE LANCIATO TOPONE
                                            ' L'INTERVALLO SI ACCORCIA
                        MsgBox msg(2), vbInformation, msg(16)
                        allerta_conn = True
                        Shell "topone.exe " & altercode, vbNormalFocus
                        Timer1.Interval = 5000
                             
                                Case Else
                                'End
                               End Select
                    
                    End If
                    
   End Select
    
  

 
End Sub



