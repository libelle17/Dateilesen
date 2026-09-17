Attribute VB_Name = "Module3"
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

'get os version
Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Sub Sendkeys(Text$, Optional wait As Boolean = False)

    Static init As Boolean, IsIDEUnderVista As Boolean, WshShell As Object
    
    'wrapper for Sendkeys which does not cause an Error 70 in the IDE under Windows Vista/ 7
    'WshShell does not get set to Nothing before App ends but any problems related to that only happen in the IDE
    'Errors due to WScript disablement on the OS can only happen under the IDE
    'Extra overhead for the WshShell object is only required under the IDE

    If Not init Then
        If IsDevEnv() Then
            IsIDEUnderVista = (OsVersion() >= 6)
            If IsIDEUnderVista Then Set WshShell = CreateObject("WScript.Shell")
        End If
        init = True
    End If
    
    If Not IsIDEUnderVista Then
        VBA.Sendkeys Text$, wait
    Else
        WshShell.Sendkeys Text$, wait
    End If
    
End Sub

Public Function IsDevEnv() As Boolean

    Dim strFileName$
    Dim lngCount&
    
    strFileName = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFileName, 255&)
    strFileName = Left(strFileName, lngCount)
    
    IsDevEnv = (UCase(Right(strFileName, 7)) Like "VB?.EXE")

End Function

Function OsVersion() As Single
    
    Dim os As OSVERSIONINFO
    Dim RetVal As Long
    os.dwOSVersionInfoSize = Len(os)
    RetVal = GetVersionEx(os)  ' read Windows's version information
    
'    MsgBox "Windows version number:" & os.dwMajorVersion & Chr(Asc(".")) & os.dwMinorVersion
'    MsgBox "OS Version Info Size = " & os.dwOSVersionInfoSize
'    MsgBox "BuildNumber = " & os.dwBuildNumber
'    MsgBox "Platform ID = " & os.dwPlatformID 'Note If ID =0 win 3.x , ID=1 win9x and ID =2 WINNT
'    MsgBox "CSD Version = " & os.szCSDVersion
    
    
    OsVersion = Val(os.dwMajorVersion & "." & os.dwMinorVersion)

End Function

