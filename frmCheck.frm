VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmCheck 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche mises à jour en ligne"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   8280
   Icon            =   "frmCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   360
      Top             =   2760
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
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
      RecordSource    =   "azienda"
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
   Begin VB.Label lblattiva 
      Caption         =   "lblattiva"
      DataField       =   "Pagamento"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "lingua"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "riga3"
      DataSource      =   "Adodc1"
      Height          =   615
      Index           =   5
      Left            =   3600
      TabIndex        =   9
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "riga2"
      DataSource      =   "Adodc1"
      Height          =   615
      Index           =   4
      Left            =   3600
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "registrazione"
      DataField       =   "Provincia"
      DataSource      =   "Adodc1"
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   7
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "Indirizzo"
      DataSource      =   "Adodc1"
      Height          =   615
      Index           =   1
      Left            =   4440
      TabIndex        =   6
      Top             =   2280
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      DataField       =   "NumeroMassimo"
      DataSource      =   "Adodc2"
      Height          =   495
      Left            =   6840
      TabIndex        =   5
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      DataField       =   "Azienda"
      DataSource      =   "Adodc1"
      Height          =   615
      Index           =   0
      Left            =   3720
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      Caption         =   "Attendre. Recherche mises à jour. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label lblSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Puddy Davidson - http://puddys-world.com

'This is my example of a supreme auto updater that handles the following:
    'can disable the program from future use, and explains why
    'checks for new versions and gives an option to download it, including update release notes
    'deletes the old version from users computer, and automatically opens the new version

'I give this to you all freely, and all i ask is for you to vote for it if you like it, comments would be nice also
'use of this example is given freely to all, you may modify it as you wish
'Again its the first time i have asked ppl to vote, so if you use it thats all i ask

'Example:
'I have compiled this exact source into version 1.5 and i have uploaded to my server called UpdatecoiffurefacileNEW.exe
'to test this auto updater as is with no changes, compile this project (version is set at 1.0)
'run this project, you will see there is a new version available (1.5 as is noted on my server)
'Enjoy guys :)


'declare some variables that we will use
Dim size As Long, Remaining As Long, NowSize As Long
Dim ProgressReal As Integer, Chunk() As Byte
Dim FileName As String, updateText As String, contentsOf As String
Dim reason As String, newVer As String, curVer As String, updateString As String


Private Sub Form_Load()
'On Error GoTo ErrHandler 'add simple error handling to catch any runtime error
On Error Resume Next

lingua = Label1
'If one instance allready running, close down
'If App.PrevInstance Then End

   lingue
   
 Dim st()    As Byte     'Message compressé
    Dim oHexa   As String   'Forme Hexa du message
    Dim oAscii  As String   'Forme Ascii
    Dim oDest   As String   'Message décompressé
    
    Dim sMsg    As String
    Dim I       As Long
        
      Label3 = frmOrdini.Text1
    frmOrdini.Hide
   oDest = "Licenza: " & Label2(0) & " " & Label2(1) & " " & Label2(2) & " " & Label3  ' & " " & Label2(4) & " " & Label2(5)
   
'AGGIORNA IL NUMERO DI ORDINE SULLA TABELLA AZIENDA
'Adodc1.Recordset!massimo = Val(Label3)
'Adodc1.Recordset.Update

    
'lingua = "1F"
Select Case lingua
    Case "1F"
      codicelingua = "fr"
      Case "2I"
      codicelingua = "it"
        Case Else
        codicelingua = "en"
        End Select
      
'If Dir(App.Path & "\topone.exe") <> "" Then
'MsgBox "File Exists!"

'Else
'MsgBox "File Does Not Exist!"
'Form3.Hide

'End If

'CONTATORE DEGLI ACCESSI.... SPARTANO
'***************************************************************************************
' updateString = Inet1.OpenURL("http://axasoft.altervista.org/topone/counter/clic_1.php3?lingua=" & Label1 & Chr$(38) & "city=" & oDest & Chr$(38) & "ticket=" & Label3)
'***************************************************************************************

 Select Case Len(updateString)
 Case Is = 0
       '   MsgBox MSG(2), vbInformation, MSG(16)
         '     allerta_conn = True
         '   End
            
        Case Is > 100
       '  MsgBox MSG(17), vbInformation, MSG(16)
        '      allerta_conn = True
         '  End
        ' Case Else
         End Select
 
 altercode = serialesoftware & "#" & "PointVente" & "#" & codicelingua & "#" & Label2(0) & "#" & Label2(1) & "#" & Label2(2)
'MsgBox altercode

 
 
 
 Check3
 
 If isRegistered = False Then
 'MsgBox isRegistered
 Shell "topone.exe " & altercode, vbNormalFocus
 End If
    
    Exit Sub
    
    
    
        'run this to clean up after the updater ran (if it ran, doesnt matter)
   
    
        'for this example we will be using only the major and minor of the programs internal version, you can add revision to your own projects/server
        'curVer is the string that holds the version of the program that is running this updater
        curVer = App.Major & "." & App.Minor '& "." & App.Revision
        
        'the most current version gets put into a txt document and uploaded here
        updateString = Inet1.OpenURL("http://axasoft.altervista.org/topone/download/newVersion.txt")
    
     'newVer is the string that holds the most current version number
    If updateString <> "" Then
        newVer = updateString
    End If
    
     'if new version is higher than the running version
    If newVer > curVer Then
    
         'update notes describe what updates have been done (optional)
        updateText = Inet1.OpenURL("http://axasoft.altervista.org/topone/download/updateText" & lingua & ".txt")
         MsgBox updateText, , messaggi(8)
         Form3.Show
         End If
         
ErrHandler:
       ' MsgBox Err.Description 'if there was an error, get a description and open the main program, debug purpose mostly
       ' frmMain.Show
       ' Unload Me
End Sub

Private Sub cmdCancel_Click()
         'cancel the update
        frmCheck.Tag = "Cancel"
End Sub

Private Sub deleteOLD()

        'create the batch file in the same directory as the old and new versions to make this batch smaller
    Open App.Path & IIf(Right(App.Path, 1) <> "\", "\DeleteOLD.bat", "DeleteOLD.bat") For Output As #1 'create the batch file
    
        'open the created batch file and print some commands into it, batch file will look like this
            
            '@Echo off
            ':S
            'Del "(this is the app exe name, we use this incase the user changed the exe name)"   <note: the quotation marks throughout this batch file are nesasary incase your exe name contains spaces>
            'If Exist "(app name again here)" Goto S   <so if its not deleted yet, go back to :S and read on>
            ':D
            'ren "UpdatecoiffurefacileNEW.exe" "coiffurefacile.exe"   <use the batch to change the new version into the same name as the old version>
            'If Exist "UpdatecoiffurefacileNEW.exe" Goto D   <same as three lines above>
            'Update Example   <run the new version, name is now the same as old version>
            'Del DeleteOLD.bat   <delete this batch file>
            
    Print #1, "@Echo off" & vbCrLf & _
              ":S" & vbCrLf & _
              "Del " & Chr(34) & App.EXEName & ".exe" & Chr(34) & vbCrLf & _
              "If Exist " & Chr(34) & App.EXEName & ".exe" & Chr(34) & " Goto S" & vbCrLf & _
              ":D" & vbCrLf & _
              "ren " & Chr(34) & "UpdatetoponeNEW.exe" & Chr(34) & " " & Chr(34) & "topone.exe" & Chr(34) & vbCrLf & _
              "If Exist " & Chr(34) & "UpdatetoponeNEW.exe" & Chr(34) & " Goto D" & vbCrLf & _
              Chr(34) & "topone.exe" & Chr(34) & vbCrLf & "Del DeleteOLD.bat"
    Close #1
    
         'run the batch file, make it run hidden
        Shell "DeleteOLD.bat", vbHide
            End
End Sub


Private Sub lingue()
Select Case lingua
  
 Case Is = "1F"
 frmCheck.Caption = "Recherche mises à jour en ligne"
 lblAction.Caption = "Attendre. Recherche mises à jour. "
 cmdCancel.Caption = "Annuller"
 
msg(1) = "Attendre, Autentication!"
msg(2) = "Attention il existe un problème de connexion d'Internet. Nous vous invitons à résoudre le problème afin de déverrouiller le logiciel de gestion "
msg(3) = "Reason being:"
msg(4) = "Télécharger la mise à jour? Oui or No?"
msg(5) = "En ligne est une nouvelle version du programme ver. "
msg(6) = "C'est la ver. - "
msg(7) = "Vous souhaitez télécharger la nouvelle version?"
msg(8) = "Mise à jour disponible"
msg(9) = "Mise  jour en cours. Attendre!"
msg(10) = "Annuller"
msg(11) = "Mise  jour annullé"
msg(12) = "La mise à jour a été effectuée. Appuyez sur OK pour lancer le programme."
msg(13) = "Mise à jour complète"
msg(14) = "Si vous choisissez Non, vous serez invité à mettre à jour plus tard."
msg(15) = " - Téléchargè"
msg(16) = "Attention"
msg(17) = "Attention il existe un problème. Nous vous invitons à contacter l'editeur du logiciel."
msg(18) = "Notices de l'editeur du logiciel"
 Case Is = "2I"
 frmCheck.Caption = "Ricerca aggiornamenti online"
 lblAction.Caption = "Ricerca aggiornamenti in corso. Attendere! "
 cmdCancel.Caption = "Annullare"
 
 
 msg(1) = "Verifica in corso.Attendere"
msg(2) = "Attenzione, la connessione ad internet non risulta attiva. Sei pregato di risolvere il problema e riavviare "
msg(3) = "Il motivo può essere:"
msg(4) = "Scaricare l'aggiornamento? Si o No?"
msg(5) = "In linea c'è una nuova versione del programma. ver. "
msg(6) = "E' la ver. - "
msg(7) = "Vuoi scaricare la nuova versione?"
msg(8) = "E' disponibile un aggiornamento"
msg(9) = "Aggiornamento in corso. Attendere!"
msg(10) = "Annullare"
msg(11) = "Aggiornamento annullato"
msg(12) = "L'aggiornamento è stato effettuato. PremerE ok per avviare il programma."
msg(13) = "Aggiornamento completato"
msg(14) = "Se scegliete NO, Vi sarà proposto l'aggiornamento la prossima volta."
msg(15) = " - Scaricato"
msg(16) = "Attenzione"
msg(17) = "Attenzione, è stato rilevato un problema. Sei pregato di contattare l'autore del software."
msg(18) = "Informazioni dell'autore del software"

Case Else
frmCheck.Caption = "Check For Update "
 lblAction.Caption = "Authenticating, please wait "
 cmdCancel.Caption = "Cancel"
 
 msg(1) = "Authenticating, please wait"
msg(2) = "Internet connection inactive. Please, connect it and run again."
msg(3) = "Reason being:"
msg(4) = "Download update? Yes or No"
msg(5) = "There is a new version available v "
msg(6) = "This update - "
msg(7) = "Would you like to download the new version?"
msg(8) = "Update available"
msg(9) = "Downloading update, please wait"
msg(10) = "Cancel"
msg(11) = "Update aborted"
msg(12) = "The update is complete. Press ok to open the new version."
msg(13) = "Update Complete"
msg(14) = "You choose not to update this time, you will be asked again next time you open this program."
msg(15) = " - Downloaded"
msg(16) = "Warning"
msg(17) = "Warning. There is a problem. Please, contact the software house."
msg(18) = "Informations of software house"

 End Select
 
End Sub
Private Sub Check3() 'VERIFICA SE E' UNA COPIA REGISTRATA


Dim codicesoftware, updateString As String
Dim rilevazione As Integer

updateString = Inet1.OpenURL("http://axasoft.altervista.org/banner/software/users/registrati/" & serialesoftware & ".txt")
rilevazione = Val(updateString) 'VALORE DELLA STRINGA SCARICATA
'MsgBox rilevazione

codicesoftware = Left(serialesoftware, 5) & _
                   Right(serialesoftware, 1) & _
                   Left(Right(serialesoftware, 2), 1) & _
                     Left(Right(serialesoftware, 3), 1) & _
                      Trim(99999 - Val(Right(serialesoftware, 5)))

'MsgBox Len(lblattiva), , Len(codicesoftware)
'updateString = ""
Select Case Len(updateString)
'RILEVIAMO DELLE INFORMAZIONI DALLA LUNGHEZZA DELLA STRINGA SE NON RIUSCIAMO AD OTTENERE UN VALORE NUMERICO
   ' NEL CASO DI VALORE NUMERICO CONFRONTIAMO TRA IL VALORE RILEVATO E IL VALORE DEL CONTATORE

      Case Is = 0 ' CONNESSIONE ASSENTE: allerta_conn=true
      '            '
      '  CONTROLLARE SE E' REGISTRATA
      'allerta_conn = True
      If Left(lblattiva, 13) = Left(codicesoftware, 13) Then
         isRegistered = True
        ' MsgBox isRegistered
         
        
         
          Else: isRegistered = False
         End If
     ' MsgBox lblattiva , , Trim(codicesoftware)
      
       Exit Sub
            
        Case Is > 100  'NON ESISTE LA PAGINA ONLINE in questo caso si riceve una pagina vuota
                        ' IL PROGRAMMA SI FERMA
                        
              'SOFTWARE NON REGISTRATO
            
             isRegistered = False
             
              
              
           
           Case Is > 10 ' IN QUESTO CASO SI VALUTA UN MESSAGGIO DELL'EDITORE CHE SIA INFERIORE O UGUALE A 100 CARATTERI
                        ' SE FOSSE SUPERIORE A 100 VERREBBE VISUALIZZATO PRIMA
                        ' IN QUESTO CASO SI RICEVE IL CODICE DI ATTIVAZIONE
                        ' VALE SOLO PER CONNESSIONE ATTIVA
                        
                        codice_attivazione = updateString
                        lblattiva = codice_attivazione
                        Adodc1.Recordset!pagamento = codice_attivazione
                        Adodc1.Recordset.Update
                        allerta_conn = False
                        isRegistered = True
                        
                   '  MsgBox lblattiva, , Trim(codicesoftware) & Len(Trim(codicesoftware))
                        Exit Sub
               



               
             
             Case Else
                 
              
                    
   End Select
    
   'IN CASO DI PROBLEMI VIENE RIMOSSO IL CODICE DI ATTIVAZIONE DAL DATABASE
                  'IMMETTENDO IL VALORE 100 SUL FILE DI REGISTRAZIONE
                  
              If rilevazione = 100 Then
              
              lblattiva = " "
              isRegistered = False
             
              Adodc1.Recordset!pagamento = lblattiva
              Adodc1.Recordset.Update
              
             ' MsgBox "il codice attivazione: " & lblattiva, , "eliminato"
              End If



End Sub

