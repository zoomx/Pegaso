Attribute VB_Name = "ApiFunctions"
Option Explicit
'Sezione DECLARE
'Routine di ritardo in millisecondi
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Ritorna lo spazio libero su disco
Declare Function GetDiskFreeSpace Lib "kernel32" Alias _
"GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters _
As Long) As Long
'Ok = GetDiskFreeSpace("c:\", SectorsPerCluster, _
BytesPerSector, NumberOfFreeClusters, _
TtoalNumberOfClusters)
'Bytes = NumberOfFreeClusters * SectorsPerCluster * BytesPerSector
'Clusters = NumberOfFreeClusters * SectorsPerCluster

Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As _
String, ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, _
     ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
     Public Orario As String

'Costants for ShellExecute
Public Const SW_SHOWNORMAL = 1


Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName _
As String, ByVal lpKeyName As String, ByVal lpDefault As _
String, ByVal lpReturnedString As String, ByVal nSize As _
Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" _
Alias "WritePrivateProfileStringA" (ByVal _
lpApplicationName As String, ByVal lpKeyName As String, _
ByVal lpString As String, ByVal lpFileName As String) As Long


Function sReadINI(AppName, KeyName, filename As String) As String
'*Returns a string from an INI file. To use, call the  *
'*functions and pass it the AppName, KeyName and INI   *
'*File Name, [sReg=sReadINI(App1,Key1,INIFile)]. If you *
'*need the returned value to be a integer then use the *
'*val command.                                         *
'*******************************************************

Dim sRet As String
    sRet = String(255, Chr(0))
    sReadINI = Left(sRet, GetPrivateProfileString(AppName, ByVal KeyName, "", sRet, Len(sRet), filename))
End Function

Public Function WriteINI(sAppname, sKeyName, sNewString, sFileName As String) As Long
'*Writes a string to an INI file. To use, call the     *
'*function and pass it the sAppname, sKeyName, the New *
'*String and the INI File Name,                        *
'*[R=WriteINI(App1,Key1,sReg,INIFile)]. Returns a 1 if *
'*there were no errors and a 0 if there were errors.   *
'*******************************************************


    WriteINI = WritePrivateProfileString(sAppname, sKeyName, sNewString, sFileName)
End Function






