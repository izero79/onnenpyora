Attribute VB_Name = "Module1"
Public Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SW_SHOWNORMAL = &H40

#If Win32 Then
Public Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

#Else
Public Declare Function ShellExecute Lib _
    "shell.dll" (ByVal _
    hWnd As Integer, ByVal lpOperation _
    As String, ByVal lpFile _
    As String, ByVal lpParameters _
    As String, ByVal lpDirectory _
    As String, ByVal nShowCmd _
    As Integer) As Integer
#End If


