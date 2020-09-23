Attribute VB_Name = "modNoFocus"
Option Explicit

'API Declarations
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'Consts
Private Const GWL_WNDPROC = (-4)
Private Const WM_SETFOCUS = &H7

'Vars
Public StandardButtonProc   As Long

Public Sub NoFocusRect(Button As Object, vValue As Boolean)

    If vValue = True Then                                       'Focus rect on
        'Save the adress of the standard button procedure
        StandardButtonProc = GetWindowLong(Button.hwnd, GWL_WNDPROC)
        'Subclass the button to control its Windows Messages
        SetWindowLong Button.hwnd, GWL_WNDPROC, AddressOf ButtonProc
    Else                                                        'Focus rect off
        'Remove the subclassing from the button
        SetWindowLong Button.hwnd, GWL_WNDPROC, StandardButtonProc
    End If

End Sub

Public Function ButtonProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next

    'The procedure that gets all windows messages for the subclassed button

    Select Case uMsg&
            'The button is going to get the focus
        Case WM_SETFOCUS
            'Exit the procedure -> The message doesn´t reach the button
            Exit Function
    End Select

    'Call the standard Button Procedure
    ButtonProc = CallWindowProc(StandardButtonProc, hwnd&, uMsg&, wParam&, lParam&)

End Function

