VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOpenSaveDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OpenFileName) As Long

Private Type OpenFileName
    lStructSize         As Long
    hwndOwner           As Long
    hInstance           As Long
    lpstrFilter         As String
    lpstrCustomFilter   As String
    nMaxCustFilter      As Long
    nFilterIndex        As Long
    lpstrFile           As String
    nMaxFile            As Long
    lpstrFileTitle      As String
    nMaxFileTitle       As Long
    lpstrInitialDir     As String
    lpstrTitle          As String
    flags               As Long
    nFileOffset         As Integer
    nFileExtension      As Integer
    lpstrDefExt         As String
    lCustData           As Long
    lpfnHook            As Long
    lpTemplateName      As String
End Type

Public mFileName        As String
Private mStatus         As Boolean

Public Property Get GetName() As String

    GetName = mFileName                                         'Return the filename

End Property

Public Property Let GetName(FName As String)

    mFileName = FName                                           'Store the filename

End Property

Public Property Get GetStatus() As Boolean

    GetStatus = mStatus                                         'Return the status

End Property

Public Property Let GetStatus(FStatus As Boolean)

    mStatus = FStatus                                           'Store the status

End Property


Public Function OpenFileDialog(FormName As Long, AppInstance As Long, InitDir As String, FileFilter As String) As Long
    On Error Resume Next

    Dim X           As Long
    Dim sFilter     As String
    Dim OpenFile    As OpenFileName

    With OpenFile
        .lStructSize = Len(OpenFile)
        .hwndOwner = FormName
        .hInstance = App.hInstance
        .lpstrFilter = FileFilter
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(OpenFile.lpstrFile) - 1
        .lpstrFileTitle = OpenFile.lpstrFile
        .nMaxFileTitle = OpenFile.nMaxFile
        .lpstrInitialDir = InitDir
        .lpstrTitle = "Select Access Database to Use."
        .flags = 0
    End With

    'Ask user to enter filename to be saved
    X = GetSaveFileName(OpenFile)

    If X = 0 Then
        mFileName = vbNullString
        mStatus = False
    Else
        mStatus = True
        mFileName = Trim(OpenFile.lpstrFile)
    End If

End Function
