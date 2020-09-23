VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Identify MS Access Users"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6510
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFile 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   5835
   End
   Begin VB.CommandButton cmdOpenDialog 
      Height          =   315
      Left            =   6000
      Picture         =   "frmMain.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Get User Information"
      Top             =   180
      Width           =   315
   End
   Begin MSFlexGridLib.MSFlexGrid grdDisplay 
      Height          =   2895
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5106
      _Version        =   393216
      FixedCols       =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_GetUserInfo 
         Caption         =   "&Get User Information"
      End
      Begin VB.Menu mnuFile_Spacer1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuFile_Help 
      Caption         =   "Help"
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsOpenDialog   As New cOpenSaveDialog

'-----------------------------------------------------------------------
' Sub OpenJetUserInfo
'
' Created by Mark Bader
' Date: 12-17-2003
'
' Purpose: Gets information of users logged into the selected mdb
'
'    strDbPath: Path and File name of mdb
'
'-----------------------------------------------------------------------
Sub OpenJetUserInfo(strDbPath As String)
    On Error GoTo Proc_Err
    
    Dim cnn         As ADODB.Connection
    Dim rst         As ADODB.Recordset
    Dim fld         As ADODB.Field
    Dim strConnect  As String
    Dim strFld1     As String

    ' Change cursor to hourglass
    Me.MousePointer = vbHourglass

    ' Format connection string to open database.
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDbPath

    Set cnn = New ADODB.Connection
    cnn.Open strConnect

    ' Open user information schema query.
    Set rst = cnn.OpenSchema(Schema:=adSchemaProviderSpecific, _
            SchemaID:="{947bb102-5d43-11d1-bdbf-00c04fb92675}")

    ' Print user information to the Immediate pane.
    With rst

        Do Until .EOF

            For Each fld In .Fields
                If Not IsNull(fld.Value) Then
                    strFld1 = strFld1 & GetValidString(fld.Value) & vbTab
                End If
            Next fld

            grdDisplay.AddItem strFld1
            strFld1 = vbNullString
            .MoveNext

        Loop

    End With

    grdDisplay.FixedRows = 1

Proc_Exit:
    Me.MousePointer = vbDefault     ' Change cursor to default
    Exit Sub
    
Proc_Err:
    MsgBox "Error #" & Err.Number & vbNewLine & Err.Description, vbInformation
    Resume Proc_Exit

End Sub


'-----------------------------------------------------------------------
' Sub cmdOpenDialog_Click
'
' Created by Mark Bader
' Date: 12-17-2003
'
' Purpose: Opens the dialog box for user to select mdb file.
'
'-----------------------------------------------------------------------
Private Sub cmdOpenDialog_Click()

    Dim strFilter   As String

    strFilter = "MS Access Files (*.mdb)" & Chr(0) & "*.mdb" & Chr(0)

    Call SetupGrid
        
    If txtFile = vbNullString Then
        Call clsOpenDialog.OpenFileDialog(Me.hwnd, App.hInstance, "C:\", strFilter)
    Else
        Call clsOpenDialog.OpenFileDialog(Me.hwnd, App.hInstance, txtFile, strFilter)
    End If
    
    txtFile = clsOpenDialog.GetName

    If txtFile <> vbNullString Then
        Call OpenJetUserInfo(txtFile)
    End If

End Sub

Private Sub Form_Load()

    Dim clsOpenDialog As New cOpenSaveDialog                    'Create object

    Call NoFocusRect(Me.cmdOpenDialog, True)                    'Eliminate Focus Rectangle

    Call SetupGrid                                              'Setups the grid for display

End Sub


'-----------------------------------------------------------------------
' Sub SetupGrid
'
' Created by Mark Bader
' Date: 12-17-2003
'
' Purpose: Sets up the grid to be used
'
'-----------------------------------------------------------------------
Private Sub SetupGrid()

    With grdDisplay

        .Rows = 0                                               'Set number of rows
        .Cols = 4                                               'Set number of columns
        .ColWidth(0) = 2000                                     'Hide Column holding key
        .ColWidth(1) = 2000                                     'Set Software Title width
        .ColWidth(2) = 1100                                     'Set Version Width
        .ColWidth(3) = 850                                      'Set Company width

        .Redraw = False                                         'Turn off to reduce Flicker

        'Add Header Text
        .AddItem "Computer Name" & vbTab & _
                "Login Name" & vbTab & _
                "Is Connected" & vbTab & _
                "Bad State"

        .Redraw = True

    End With

End Sub


'-----------------------------------------------------------------------
' Function GetValidString
'
' Created by Mark Bader
' Date: 12-17-2003
'
' Purpose: Remove all invalid characters from string
'
'    strFld:        Input string
'    Return value:  Output string
'
'-----------------------------------------------------------------------
Private Function GetValidString(strFld As String) As String

    Dim iCntr   As Integer
    Dim strData As String

    For iCntr = 1 To Len(strFld)
        Select Case Asc(Mid(strFld, iCntr, 1))
            Case 48 To 122
                strData = strData & Mid(strFld, iCntr, 1)
            Case Else
        End Select
    Next iCntr

    GetValidString = strData

End Function

Private Sub Form_Unload(Cancel As Integer)

    Set clsOpenDialog = Nothing

End Sub

Private Sub mnuFile_Exit_Click()

    Unload Me
    
End Sub

Private Sub mnuFile_GetUserInfo_Click()

    Call cmdOpenDialog_Click
    
End Sub

Private Sub mnuHelp_About_Click()

    ShellAbout Me.hwnd, App.Title, "Created by MDJ Systems", ByVal 0&
    
End Sub
