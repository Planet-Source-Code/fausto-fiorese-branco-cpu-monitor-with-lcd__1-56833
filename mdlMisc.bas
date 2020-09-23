Attribute VB_Name = "mdMisc"

Option Explicit
Const SPACE = 5
Const BAR_WIDTH = 50
Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&
Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2&
Public Const SWP_NOACTIVATE = &H10&
Public Const SWP_SHOWWINDOW = &H40&
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Private Type MEMORYSTATUS
        dwLength As Long
        dwMemoryLoad As Long
        dwTotalPhys As Long
        dwAvailPhys As Long
        dwTotalPageFile As Long
        dwAvailPageFile As Long
        dwTotalVirtual As Long
        dwAvailVirtual As Long
End Type
Public memoryInfo As MEMORYSTATUS
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public GraphPoints(0 To 99) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Function DiscSpace(ByVal DrvLetter As String, ByVal intOption As Integer)
    Dim Status As Long
    Dim TotalBytes As Currency
    Dim FreeBytes As Currency
    Dim BytesAvailableToCaller As Currency
    
    Status = GetDiskFreeSpaceEx(DrvLetter, BytesAvailableToCaller, TotalBytes, FreeBytes)
    
    If Status <> 0 Then
        Select Case intOption
            Case 0: DiscSpace = FreeBytes * 10000
            Case 1: DiscSpace = TotalBytes * 10000
        End Select
    End If

End Function


Public Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    
    GetVersionEx OSInfo
    
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
