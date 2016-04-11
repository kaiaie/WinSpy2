VERSION 5.00
Begin VB.Form FSpy 
   Caption         =   "Window Spy"
   ClientHeight    =   5340
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Spy.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSetRect 
      Caption         =   "&Set"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox txtRect 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   2400
      TabIndex        =   10
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRect 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   600
      TabIndex        =   8
      Text            =   "0"
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox txtRect 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtRect 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   4
      Text            =   "0"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtWindowClass 
      BackColor       =   &H8000000F&
      ForeColor       =   &H80000012&
      Height          =   285
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstTree 
      Height          =   1980
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label lblRect 
      AutoSize        =   -1  'True
      Caption         =   "&Bottom:"
      Height          =   195
      Index           =   3
      Left            =   1920
      TabIndex        =   9
      Top             =   3000
      Width           =   570
   End
   Begin VB.Label lblRect 
      AutoSize        =   -1  'True
      Caption         =   "&Right:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   435
   End
   Begin VB.Label lblRect 
      AutoSize        =   -1  'True
      Caption         =   "&Top:"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label lblRect 
      AutoSize        =   -1  'True
      Caption         =   "&Left:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   345
   End
   Begin VB.Label lblWindowClass 
      AutoSize        =   -1  'True
      Caption         =   "Window &Class:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   1050
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileRefresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FSpy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4

Private Const GW_CHILD = 5
Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2

Private moLayout As New CGridLayout

Private Const RECT_LEFT = 0
Private Const RECT_TOP = 1
Private Const RECT_RIGHT = 2
Private Const RECT_BOTTOM = 3

Private maFocusHandlers As New Collection

Private Sub FillList()
   Dim hWnd As Long
   Dim lLen As String
   Dim sTitle As String
   Dim l As Long
   With Me.lstTree
      .Clear
      hWnd = Me.hWnd
      hWnd = GetWindow(hWnd, GW_HWNDFIRST)
      Do While hWnd <> 0
         lLen = GetWindowTextLength(hWnd)
         If lLen > 0 Then
            sTitle = VBA.String(lLen + 1, VBA.Chr(0))
            l = GetWindowText(hWnd, sTitle, lLen + 1)
            If l <> 0 Then
               If VBA.Right(sTitle, 1) = VBA.Chr(0) Then
                  sTitle = VBA.Left(sTitle, VBA.Len(sTitle) - 1)
               End If
               sTitle = VBA.CStr(hWnd) & " - " & sTitle
               .AddItem sTitle
               .ItemData(.NewIndex) = hWnd
            End If
         End If
         hWnd = GetWindow(hWnd, GW_HWNDNEXT)
      Loop
   End With
End Sub


Private Sub GetKids(ByVal ListIndex As Long)
   Dim hWnd As Long
   Dim lLen As String
   Dim sTitle As String
   Dim sIndent As String
   Dim l As Long
   
   sIndent = VBA.String(GetLevel(ListIndex) + 1, vbTab)
   With Me.lstTree
      hWnd = .ItemData(ListIndex)
      hWnd = GetWindow(hWnd, GW_CHILD)
      Do While hWnd <> 0
         lLen = GetWindowTextLength(hWnd)
         If lLen > 0 Then
            sTitle = VBA.String(lLen + 1, VBA.Chr(0))
            l = GetWindowText(hWnd, sTitle, lLen + 1)
            If l <> 0 Then
               If VBA.Right(sTitle, 1) = VBA.Chr(0) Then
                  sTitle = VBA.Left(sTitle, VBA.Len(sTitle) - 1)
               End If
            Else
               sTitle = "[Untitled window]"
            End If
         Else
            sTitle = "[Untitled window]"
         End If
         sTitle = sIndent & VBA.CStr(hWnd) & " - " & sTitle
         .AddItem sTitle, ListIndex + 1
         .ItemData(.ListCount - 1) = hWnd
         hWnd = GetWindow(hWnd, GW_HWNDNEXT)
      Loop
   End With
End Sub

Private Function GetLevel(ByVal ListIndex As Long) As Long
   Dim lTabCount As Long
   Dim I As Long
   Dim l As Long
   Dim sText As String
   
   sText = Me.lstTree.List(ListIndex)
   l = VBA.Len(sText)
   For I = 1 To l
      If VBA.Mid(sText, I, 1) <> vbTab Then Exit For
      lTabCount = lTabCount + 1
   Next I

   GetLevel = lTabCount
End Function

Private Sub cmdSetRect_Click()
    Dim hWnd As Long
    
    hWnd = lstTree.ItemData(lstTree.ListIndex)
    SetWindowPos _
        hWnd, _
        0, _
        CLng(txtRect(RECT_LEFT).Text), _
        CLng(txtRect(RECT_TOP).Text), _
        CLng(txtRect(RECT_RIGHT).Text) - CLng(txtRect(RECT_LEFT).Text), _
        CLng(txtRect(RECT_BOTTOM).Text) - CLng(txtRect(RECT_TOP).Text), _
        SWP_NOACTIVATE Or SWP_NOZORDER
End Sub

Private Sub Form_Load()
    Dim oFocusHandler As KTxtFocus
    ' Set text focus handlers
    Set oFocusHandler = New KTxtFocus
    Set oFocusHandler.Text = txtWindowClass
    maFocusHandlers.Add oFocusHandler
    ' Initialise layout
    With moLayout
        .Initialize 11, 9
        .RowHeight(1) = 120
        .RowHeight(2) = "100%"
        .RowHeight(3) = 120
        .RowHeight(4) = txtWindowClass.Height
        .RowHeight(5) = 120
        .RowHeight(6) = txtRect(RECT_LEFT).Height
        .RowHeight(7) = 120
        .RowHeight(8) = txtRect(RECT_RIGHT).Height
        .RowHeight(9) = 120
        .RowHeight(10) = cmdSetRect.Height
        .RowHeight(11) = 120
        .ColWidth(1) = 120
        .ColWidth(2) = MUtils.MaxOf(lblWindowClass.Width, lblRect(RECT_LEFT).Width, lblRect(RECT_RIGHT).Width)
        .ColWidth(3) = 120
        .ColWidth(4) = "50%"
        .ColWidth(5) = 120
        .ColWidth(6) = MUtils.MaxOf(lblRect(RECT_TOP).Width, lblRect(RECT_BOTTOM).Width)
        .ColWidth(7) = 120
        .ColWidth(8) = "50%"
        .ColWidth(9) = 120
        .SetCell 2, 2, lstTree, 1, True, ALIGN_LEFT, 7, True, ALIGN_TOP
        .SetCell 4, 2, lblWindowClass, , , , , , ALIGN_MIDDLE
        .SetCell 4, 4, txtWindowClass, , , , 5
        .SetCell 6, 2, lblRect(RECT_LEFT)
        .SetCell 6, 4, txtRect(RECT_LEFT)
        .SetCell 6, 6, lblRect(RECT_TOP)
        .SetCell 6, 8, txtRect(RECT_TOP)
        .SetCell 8, 2, lblRect(RECT_RIGHT)
        .SetCell 8, 4, txtRect(RECT_RIGHT)
        .SetCell 8, 6, lblRect(RECT_BOTTOM)
        .SetCell 8, 8, txtRect(RECT_BOTTOM)
        .SetCell 10, 8, cmdSetRect
    End With
    FillList
End Sub


Private Sub lstTree_Click()
    Dim hWnd As Long
    Const nBufferSize As Integer = 255
    Dim sClassName As String
    
    hWnd = lstTree.ItemData(lstTree.ListIndex)
    ' Get window class
    sClassName = VBA.String(nBufferSize, VBA.Chr(0))
    If GetClassName(hWnd, sClassName, nBufferSize) <> 0 Then
        If VBA.InStr(sClassName, VBA.Chr(0)) > 1 Then
            sClassName = VBA.Left(sClassName, VBA.InStr(sClassName, VBA.Chr(0)))
        End If
        Me.txtWindowClass.Text = sClassName
    Else
        Me.txtWindowClass.Text = ""
    End If
    
    ' Get window co-ordinates
    Dim tRect As RECT
    If GetWindowRect(hWnd, tRect) <> 0 Then
        txtRect(RECT_LEFT).Text = CStr(tRect.Left)
        txtRect(RECT_TOP).Text = CStr(tRect.Top)
        txtRect(RECT_RIGHT).Text = CStr(tRect.Right)
        txtRect(RECT_BOTTOM).Text = CStr(tRect.Bottom)
    Else
        txtRect(RECT_LEFT).Text = "0"
        txtRect(RECT_TOP).Text = "0"
        txtRect(RECT_RIGHT).Text = "0"
        txtRect(RECT_BOTTOM).Text = "0"
    End If
    UpdateUI
End Sub

Private Sub lstTree_DblClick()
   Dim lCurrLevel As Long
   With Me.lstTree
      If .ListIndex = .ListCount - 1 Then
         GetKids .ListIndex
      ElseIf GetLevel(.ListIndex) >= GetLevel(.ListIndex + 1) Then
         GetKids .ListIndex
      Else
         lCurrLevel = GetLevel(.ListIndex + 1)
         Do
            If GetLevel(.ListIndex + 1) >= lCurrLevel Then
               .RemoveItem .ListIndex + 1
            Else
               Exit Do
            End If
         Loop
      End If
   End With
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileRefresh_Click()
    FillList
End Sub

Private Sub mnuHelpAbout_Click()
    FAbout.Show vbModal
End Sub

Private Function IsNumericAndPositive(ByVal s As String, Optional ByVal bStrictlyPositive As Boolean = False) As Boolean
    Dim bResult As Boolean
    Dim n As Long
    
    If IsNumeric(s) Then
        n = CLng(s)
        If bStrictlyPositive Then
            bResult = (n > 0)
        Else
            bResult = (n >= 0)
        End If
    End If
    IsNumericAndPositive = bResult
End Function

Private Sub txtRect_Change(Index As Integer)
    UpdateUI
End Sub

Private Sub UpdateUI()
    cmdSetRect.Enabled = _
        IsNumericAndPositive(txtRect(RECT_LEFT).Text, True) And _
        IsNumericAndPositive(txtRect(RECT_TOP).Text, True) And _
        IsNumericAndPositive(txtRect(RECT_RIGHT).Text, True) And _
        IsNumericAndPositive(txtRect(RECT_BOTTOM).Text, True)
End Sub
