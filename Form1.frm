VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Calculate Directory Sizes"
   ClientHeight    =   4350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4350
   ScaleWidth      =   9000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Stop!"
      Height          =   1155
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8955
      TabIndex        =   4
      Top             =   4310
      Width           =   8955
      Begin VB.Line Line1 
         X1              =   60
         X2              =   9000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   315
      Left            =   6960
      TabIndex        =   3
      Top             =   60
      Width           =   315
   End
   Begin MSComctlLib.ListView LV 
      Height          =   4155
      Left            =   60
      TabIndex        =   2
      Top             =   420
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   7329
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Directory Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Size (kb)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size (bytes)"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calc Directory Sizes"
      Height          =   315
      Left            =   7320
      TabIndex        =   1
      Top             =   60
      Width           =   1635
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   60
      TabIndex        =   0
      Text            =   "h:\vb code\"
      Top             =   60
      Width           =   6855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const SB_HORZ = 0
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Dim bStop As Boolean
Private Sub Command1_Click()
Dim sDirs() As String
Dim sTemp() As String
Dim X As Long
Dim y As Long
Dim z As Long
Dim lCurPos As Long
Const sUnsearched As String = "0"
Const sSearched As String = "1"
    Command1.Enabled = False
    Command2.Enabled = False
    Text1.Enabled = False
    bStop = False
    Command3.Visible = True
    LV.ListItems.Clear
    'HideListviewScrollbar
    DoEvents
    ReDim sDirs(3, 0)
    sDirs(0, 0) = sUnsearched
    sDirs(1, 0) = Trim(Text1)
    sDirs(2, 0) = FindFilesAPI(Trim(Text1), "*.*")
    lCurPos = 0
    Do
        If bStop Then
            Command3.Visible = False
            Command1.Enabled = True
            Command2.Enabled = True
            Text1.Enabled = True
            Command3.Visible = False
            Exit Sub
        End If
        Erase sTemp
        sTemp = SubFolders(sDirs(1, lCurPos))
        On Error GoTo 0
        sDirs(0, lCurPos) = sSearched
        If UBound(sTemp) >= 0 Then
            For X = 0 To UBound(sTemp) - 1
                If X < UBound(sTemp) Then
                    ReDim Preserve sDirs(3, UBound(sDirs, 2) + 1)
                End If
                y = UBound(sDirs, 2)
                sDirs(0, y) = sUnsearched
                sDirs(1, y) = sTemp(X)
                sDirs(2, y) = Format(CDbl(FindFilesAPI(sDirs(1, y), "*.*")) / 1, "#0")
            Next X
        End If
        lCurPos = lCurPos + 1
        DoEvents
    Loop Until sDirs(0, UBound(sDirs, 2)) = sSearched
    QuickSortStringsAscending sDirs, 0, UBound(sDirs, 2)
    
    For y = UBound(sDirs, 2) To 1 Step -1
        z = y - 1
        While z > 0 And OneLevelLess(sDirs(1, y)) <> sDirs(1, z)
            z = z - 1
        Wend
        If OneLevelLess(sDirs(1, y)) = sDirs(1, z) Then
            sDirs(2, z) = CDbl(sDirs(2, z)) + CDbl(sDirs(2, y))
        End If
    Next y

    For X = 0 To UBound(sDirs, 2)
        LV.ListItems.Add X + 1, , sDirs(1, X)
        LV.ListItems.Item(X + 1).SubItems(1) = Format(CDbl(sDirs(2, X)) / 1024, "###,###,###,##0.0")
        LV.ListItems.Item(X + 1).SubItems(2) = Format(sDirs(2, X), "000000000000000")
    Next X
    Erase sDirs
    DoEvents
    'HideListviewScrollbar
    Command3.Visible = False
    Command1.Enabled = True
    Command2.Enabled = True
    Text1.Enabled = True
    
End Sub
Sub HideListviewScrollbar()
    ShowScrollBar LV.hwnd, SB_HORZ, 0&
End Sub
Function OneLevelLess(sIn As String) As String
    OneLevelLess = Left(sIn, InStrRev(sIn, "\", Len(sIn) - 1, vbTextCompare))
End Function
Function SubFolders(ByVal strRootDir As String) As Variant
On Error Resume Next
Dim strSubDir As String
Dim strDelimiter As String
Dim strReturn As String

    If Right(strRootDir, 1) <> "\" Then strRootDir = strRootDir & "\"
    
    strDelimiter = ";"
    strSubDir = Dir(strRootDir, vbDirectory) 'Retrieve the first entry
    
    Do While Len(strSubDir) <> 0
        If strSubDir <> "." And strSubDir <> ".." Then
            If (GetAttr(strRootDir & strSubDir) And vbDirectory) = vbDirectory Then
                strReturn = strReturn & strRootDir & strSubDir & strDelimiter
            End If
        End If
        strSubDir = Dir 'Get Next entry
    Loop
    
    SubFolders = Split(strReturn, strDelimiter)
Exit Function
ehSubFolders:
    SubFolders = Empty
End Function
Function FindFilesAPI(Path As String, SearchStr As String)
Dim FileName As String
Dim FileCount As Integer
Dim nDir As Integer
Dim i As Integer
Dim hSearch As Long
Dim WFD As WIN32_FIND_DATA
Dim Cont As Integer
Dim FT As FILETIME
Dim ST As SYSTEMTIME
Dim DateCStr As String, DateMStr As String
    FindFilesAPI = 0
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") And _
            ((GetFileAttributes(Path & FileName) And _
            FILE_ATTRIBUTE_DIRECTORY) <> FILE_ATTRIBUTE_DIRECTORY) Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
            End If
            Cont = FindNextFile(hSearch, WFD)
        Wend
    Cont = FindClose(hSearch)
    End If

End Function

Private Sub Command2_Click()

Dim MyDir As String

MyDir = BrowseFolder(hwnd, "Select Destination Directory...")
If Len(MyDir) = 0 Then Exit Sub
Text1 = MyDir & IIf(Right(Trim(MyDir), 1) = "\", "", "\")
Command1_Click

End Sub

Private Sub Command3_Click()
    bStop = True
    
End Sub

Private Sub Form_Load()
    LV.ColumnHeaders(1).Width = ((LV.Width - 300) / 100) * 80
    LV.ColumnHeaders(2).Width = ((LV.Width - 300) / 100) * 20
    LV.ColumnHeaders(3).Width = ((LV.Width - 300) / 100) * 20
    'HideListviewScrollbar
End Sub

Private Sub LV_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim iCH As Integer
If ColumnHeader.Index = 2 Then
    iCH = 2
Else
    iCH = ColumnHeader.Index - 1
End If
LV.SortKey = iCH
If LV.SortOrder = lvwAscending Then
    LV.SortOrder = lvwDescending
Else
    LV.SortOrder = lvwAscending
End If
If LV.Sorted = False Then LV.Sorted = True

End Sub
