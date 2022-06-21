Attribute VB_Name = "Module11"


Global variabile1, variabile2, variabile3, variabile4, variabile5, variabile6 As String
Global art1, art2, art3, art5, art6, art7, art8, art9, art10, art11 As String
Global art4 As Single
Global a, b, c As Single
Global valore1, valore2, valore3 As Single
Global varia1, varia2, varia3, varia4, varia5 As Single
Global systempath As String
Global somma1, somma2, somm3 As Single
Global winproductid, messaggi(14), licenza As String
Global msg(26) As String
Global archivipagamenti, indirizzo As String
Global casa, ad, google As String
Global id, erreasoft As Integer
Global iva2 As Double
Global azienda, codiceazienda, id_cliente, pw_cliente, cittadina, int_cliente As String
Global lingua, Button, button1 As String
Global security As Integer
Global campo(5) As String
Global righe As Integer
Global chiave(80) As Integer
Global table As String
Global stringa3, stringa4 As String
Global spia As Integer
Global datarow As Integer
Global colori(8) As String
Global manu As Boolean
Global Indice As Integer
Global I As Integer
Global txtFields(61) As String
Global altercode, serialesoftware, codicelingua, codice_attivazione As String
Global contatore As Integer
Global allerta_conn, isRegistered As Boolean

'INI Read and Write
Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
    As String, lpKeyName As Any, ByVal lpDefault As String, _
    ByVal lpRetunedString As String, ByVal nSize As Long, _
    ByVal lpFileName As String) As Long


Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName _
    As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
    ByVal lplFileName As String) As Long
'INI Read and Write
                        
Public Function INIRead(iAppName As String, iKeyName As String, iFileName As String) As String
    'Example:
    'x = INIRead("boot", "shell", "C:\WINDOWS\system.ini")
    ' ou boot est la section, shell la cle
    
    Dim iStr As String
    
    iStr = String(255, Chr(0))
    INIRead = Left(iStr, GetPrivateProfileString(iAppName, ByVal iKeyName, "", iStr, Len(iStr), iFileName))
    
End Function

Public Function INIWrite(iAppName As String, iKeyName As String, iKeyString As String, iFileName As String)
    'Example:
    'x = INIWrite("boot", "shell", "Explorer.exe", "C:\WINDOWS\system.ini")
    ' ou boot est la section, shell la cle
    
    r% = WritePrivateProfileString(iAppName, iKeyName, iKeyString, iFileName)
    
End Function


Public Function GetKeyVal(ByVal Section As String, ByVal Key As String, ByVal INIFileLoc As String)
    'This Function retrieves information fro
    '     m an INI File
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = Section where the Key is held
    '
    'Key = The Key of which you want to retr
    '     ieve information
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'GetKeyVal'", vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to Get
    '     Key Value
    Dim RetVal As String, Worked As Integer
    RetVal = String$(255, 0)
    Worked = GetPrivateProfileString(Section, Key, "", RetVal, Len(RetVal), INIFileLoc)


    If Worked = 0 Then
        GetINI = ""
    Else
        GetINI = Left(RetVal, InStr(RetVal, Chr(0)) - 1)
    End If
End Function


Function AddToINI(ByVal Section As String, ByVal Key As String, ByVal Value As String, ByVal INIFileLoc As String)
    'This Function adds a Section, Key, or V
    '     alue to an INI file
    'Also used to CREATE NEW INI FILE
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = The name of the referred to S
    '     ection or newly created Section (ex. "Ne
    '     w Section 1")
    'Key = The name of the referred to Key o
    '     r newly created Key (ex. "New Key 1")
    'Value = The value to hold in the given
    '     Key (ex. "New Info Held")
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'AddToINI'", vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to Add
    '     the information to the INI File
    WritePrivateProfileString Section, Key, Value, INIFileLoc
End Function


Public Function DeleteSection(ByVal Section As String, ByVal INIFileLoc As String)
    'This Function Deletes a specified Secti
    '     on from an INI file
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = The name of the Section you w
    '     ish to remove (ex. "Section Number 1")
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'DeleteSection'", vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to dele
    '     te Section
    WritePrivateProfileString Section, vbNullString, vbNullString, INIFileLoc
    'NOTE: vbNullString is the coding in whi
    '     ch to delete a Section, or Key
End Function


Public Function DeleteKey(ByVal Section As String, ByVal Key As String, ByVal INIFileLoc As String)
    'This Function Deletes a Key in a specif
    '     ied Section from an INI file
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = The name of the Section in wh
    '     ich the Key to be deleted is held (ex. "
    '     Section Number 1")
    'Key = The name of the Key you wish to r
    '     emove (ex. "Key Number 5")
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'DeleteKey'", vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to dele
    '     te Key
    WritePrivateProfileString Section, Key, vbNullString, INIFileLoc
    'NOTE: vbNullString is the coding in whi
    '     ch to delete a Section, or Key
End Function


Public Function DeleteKeyValue(ByVal Section As String, ByVal Key As String, ByVal INIFileLoc As String)
    'This Function deletes the value in a sp
    '     ecified Key from an INI file
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = The name of the Section in wh
    '     ich the Key is held (ex. "Section Number
    '     1")
    'Key = The name of the Key you wish to r
    '     emove the value from (ex. "Key Number 5"
    '     )
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'DeleteKeyValue'", vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to dele
    '     te Key Value
    WritePrivateProfileString Section, Key, "", INIFileLoc
    ' "" = is a short way of saying Nothing
End Function


Public Function RenameSection()
    'Coming Soon
End Function


Public Function RenameKey()
    'Coming Soon
End Function


Public Sub WhiteLine(iFileName As String)
'insert a white in the ini file
    Open iFileName For Append Access Read Write As #1
    Print #1, " "
    Close #1
        
End Sub
Sub Main()

End Sub


