<div align="center">

## Associate File Name and Icon


</div>

### Description

For those of you who want to add a touch of professionalism to your program, now you can create a file type in the Windows Registry database which will associate all files ending with your program's file extension ( yourfile.xxx) to your program. You also specify an icon for your file type and a description. This example also shows you how to use Command$ to open these files in your program when once the file is clicked or opened, and a quick tip on creating files in the Windows Recent file folder (Start > Documents).
 
### More Info
 


You need an icon in your program's directory which will be referenced to for file association. Open a new project, add a form (Form1) and a Module (Module1)

To use the program, run it. Now compile the program and make a file named "test.xyz" in Notepad and save it. Now click on that file named "test.xyz". Your program will open.

You need an icon in your program's directory which will be referenced to for file association.

None that I know of. If you have problems getting this to work, e-mail support@rgsoftware.com


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[RGSoftware](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rgsoftware.md)
**Level**          |Unknown
**User Rating**    |5.0 (45 globes from 9 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rgsoftware-associate-file-name-and-icon__1-1233/archive/master.zip)

### API Declarations

```
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4
```


### Source Code

```
'In a module:
'-----------------------------------------
Public Sub savekey(Hkey As Long, strPath As String)
Dim keyhand&
r = RegCreateKey(Hkey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub
Public Function getstring(Hkey As Long, strPath As String, strValue As String)
Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(Hkey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
  strBuf = String(lDataBufSize, " ")
  lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
  If lResult = ERROR_SUCCESS Then
    intZeroPos = InStr(strBuf, Chr$(0))
    If intZeroPos > 0 Then
      getstring = Left$(strBuf, intZeroPos - 1)
    Else
      getstring = strBuf
    End If
  End If
End If
End Function
Public Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub
Function getdword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long
r = RegOpenKey(Hkey, strPath, keyhand)
lDataBufSize = 4
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)
If lResult = ERROR_SUCCESS Then
  If lValueType = REG_DWORD Then
    getdword = lBuf
  End If
End If
r = RegCloseKey(keyhand)
End Function
Function SaveDword(ByVal Hkey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
  Dim lResult As Long
  Dim keyhand As Long
  Dim r As Long
  r = RegCreateKey(Hkey, strPath, keyhand)
  lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
  r = RegCloseKey(keyhand)
End Function
Public Function DeleteKey(ByVal Hkey As Long, ByVal strKey As String)
Dim r As Long
r = RegDeleteKey(Hkey, strKey)
End Function
Public Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
r = RegOpenKey(Hkey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
r = RegCloseKey(keyhand)
End Function
'-------------------------------------------
'On a Form:
'----------------------------------------------
 Private Declare Function fCreateShellLink Lib "STKIT432.DLL" (ByVal _
    lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal _
    lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long
Private Sub Form_Load()
Dim strString As String
Dim lngDword As Long
If Command$ <> "%1" Then
Msgbox (Command$ & " is the file you need to open!"), vbInformation
 'Add to Recent file folder
    lReturn = fCreateShellLink("..\..\Recent", _
    Command$, Command$, "")
End If
'create an entry in the class key
Call savestring(HKEY_CLASSES_ROOT, "\.xyz", "", "xyzfile")
'content type
Call savestring(HKEY_CLASSES_ROOT, "\.xyz", "Content Type", "text/plain")
'name
Call savestring(HKEY_CLASSES_ROOT, "\xyzfile", "", "This is where you type the description for these files")
'edit flags
Call SaveDword(HKEY_CLASSES_ROOT, "\xyzfile", "EditFlags", "0000")
'file's icon (can be an icon file, or an icon located within a dll file)
Call savestring(HKEY_CLASSES_ROOT, "\xyzfile\DefaultIcon", "", App.Path & "\ICON.ico")
'Shell
Call savestring(HKEY_CLASSES_ROOT, "\xyzfile\Shell", "", "")
'Shell Open
Call savestring(HKEY_CLASSES_ROOT, "\xyzfile\Shell\Open", "", "")
'Shell open command
Call savestring(HKEY_CLASSES_ROOT, "\xyzfile\Shell\Open\command", "", App.Path & "\Project1.exe %1")
End Sub
'----------------------------------------------
```

