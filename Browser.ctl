VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl Browser 
   BackColor       =   &H008080FF&
   BackStyle       =   0  'Transparent
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9915
   DefaultCancel   =   -1  'True
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   4290
   ScaleWidth      =   9915
   ToolboxBitmap   =   "Browser.ctx":0000
   Begin VB.Timer tmrTick 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   3345
      Top             =   3975
   End
   Begin VB.VScrollBar scScroll 
      Enabled         =   0   'False
      Height          =   3915
      LargeChange     =   10
      Left            =   9690
      Max             =   14
      Min             =   14
      TabIndex        =   10
      Top             =   0
      Value           =   14
      Width           =   225
   End
   Begin VB.TextBox txSearch 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      ToolTipText     =   "Search for..."
      Top             =   3990
      Width           =   2115
   End
   Begin VB.CommandButton btNav 
      Caption         =   "| <"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   3780
      TabIndex        =   1
      ToolTipText     =   "First"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btNav 
      Caption         =   "> |"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   6465
      TabIndex        =   6
      ToolTipText     =   "Last"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btNav 
      Caption         =   "<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   4845
      TabIndex        =   3
      ToolTipText     =   "Previous"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btNav 
      Caption         =   ">"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   5385
      TabIndex        =   4
      ToolTipText     =   "Next"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btNav 
      Caption         =   ">>"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   5925
      TabIndex        =   5
      ToolTipText     =   "Page down"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btNav 
      Caption         =   "<<"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   4305
      TabIndex        =   2
      ToolTipText     =   "Page up"
      Top             =   3990
      Width           =   540
   End
   Begin VB.CommandButton btOKCan 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   7410
      TabIndex        =   7
      Top             =   3990
      Width           =   945
   End
   Begin VB.CommandButton btOKCan 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   8385
      TabIndex        =   8
      Top             =   3990
      Width           =   945
   End
   Begin MSFlexGridLib.MSFlexGrid gdBrowse 
      Height          =   3930
      Left            =   0
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   9690
      _ExtentX        =   17092
      _ExtentY        =   6932
      _Version        =   393216
      Rows            =   16
      Cols            =   10
      BackColorSel    =   128
      BackColorBkg    =   12632256
      GridColor       =   8421504
      GridColorFixed  =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLines       =   2
      ScrollBars      =   0
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lbFullPop 
      Caption         =   "ü"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   14.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   7095
      TabIndex        =   12
      ToolTipText     =   "Fully populated"
      Top             =   3990
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Label lbSearchName 
      Alignment       =   1  'Rechts
      Appearance      =   0  '2D
      BackColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   15
      TabIndex        =   11
      Top             =   4035
      Width           =   1095
   End
End
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
DefLng A-Z

Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Sub PutFocus Lib "user32" Alias "SetFocus" (ByVal hWnd)

Private Type ColInfo
    myColWidth          As Long
    myHeader            As String
    myFieldName         As String
    myFieldTranslation  As Boolean
End Type

Public Enum SortDirection
    Ascending = 1
    Descending = 2
End Enum

Private Const DefWidth  As Long = 510          'column 0 width
Private Const MinIx     As Long = 1
Private Const MaxIx     As Long = 15
Private Const NavCodes  As String = "$!&(""#"  'keycodes for Pos1 PageUp CursorUp CursorDown PageDown End

Private myEnabled       As Boolean
Private myAutosize      As Boolean
Private myDynamicScroll As Boolean
Private myDisplayName   As String
Private myRecordset     As DAO.Recordset
Private myOrderedBy     As String
Private mySortOrder     As SortDirection
Private myColumnInfo()  As ColInfo
Private myCurrBookmark(1 To MaxIx) As String

Private Head            As String
Private ForwardLine     As Boolean
Private ForwardPage     As Boolean
Private ReverseLine     As Boolean
Private ReversePage     As Boolean
Private NotFull         As Boolean
Private PageScroll      As Boolean
Private ScrChanged      As Boolean
Private FieldContents   As Variant
Private CompOper        As String
Private OtherCompOper   As String
Private ScrollDivi      As Long
Private TotalWidth      As Long
Private PreviousRow     As Long
Private FilledTo        As Long
Private TpP             As Long

Public Event OK()
Public Event Cancel()
Public Event PositionChanged(ByVal Row As Long)
Public Event TranslateColumn(ByVal FieldName As String, ByVal OldValue As Variant, NewValue As Variant)

Public Sub AdjustCol(ByVal Col As Long)

    ColWidth(Col) = RequiredColWidth(Col)

End Sub

Public Sub AdjustCols()

  Dim i

    For i = 1 To gdBrowse.Cols - 1
        AdjustCol (i)
    Next i

End Sub

Public Property Let Autosize(ByVal nuAutosize As Boolean)
Attribute Autosize.VB_Description = "Sets / returns whether the Control will automatically adjust the column widths to the text displayed."
Attribute Autosize.VB_HelpID = 10007

    myAutosize = (nuAutosize <> False)
    If Ambient.UserMode Then
        AdjustCols
    End If
    PropertyChanged "Autosize"

End Property

Public Property Get Autosize() As Boolean

    Autosize = myAutosize

End Property

Public Property Get Backcolor() As OLE_COLOR

    Backcolor = gdBrowse.Backcolor

End Property

Public Property Let Backcolor(ByVal nuBackColor As OLE_COLOR)

    gdBrowse.Backcolor = nuBackColor
    PropertyChanged "Backcolor"

End Property

Public Property Let BarBackcolor(ByVal nuBackColor As OLE_COLOR)
Attribute BarBackcolor.VB_Description = "Sets / returns the highlite bar backcolor."
Attribute BarBackcolor.VB_HelpID = 10012

    gdBrowse.BackColorSel = nuBackColor
    PropertyChanged "BarBackcolor"

End Property

Public Property Get BarBackcolor() As OLE_COLOR

    BarBackcolor = gdBrowse.BackColorSel

End Property

Public Property Let BarForecolor(ByVal nuForecolor As OLE_COLOR)
Attribute BarForecolor.VB_Description = "Sets / returns the highlite bar forecolor."
Attribute BarForecolor.VB_HelpID = 10013

    gdBrowse.ForeColorSel = nuForecolor
    PropertyChanged "BarForeColor"

End Property

Public Property Get BarForecolor() As OLE_COLOR

    BarForecolor = gdBrowse.ForeColorSel

End Property

Public Property Get Bookmark(Row As Long) As String

    If Row >= MinIx And Row <= MaxIx Then
        Bookmark = myCurrBookmark(Row)
    End If

End Property

Private Sub btNav_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    tmrTick.Interval = 333
    Select Case Index
      Case 0   'BOF
        GoFirst
      Case 1   'Page back
        ReversePage = True
        ScrollPageRev
        tmrTick.Enabled = True
      Case 2   'Line back
        ReverseLine = True
        ScrollLineRev
        tmrTick.Enabled = True
      Case 3   'Line forward
        ForwardLine = True
        ScrollLineFwd
        tmrTick.Enabled = True
      Case 4   'Page forward
        ForwardPage = True
        ScrollPageFwd
        tmrTick.Enabled = True
      Case 5   'EOF
        GoLast
    End Select

End Sub

Private Sub btNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    tmrTick.Enabled = False
    ForwardLine = False
    ForwardPage = False
    ReverseLine = False
    ReversePage = False
    gdBrowse.SetFocus

End Sub

Private Sub btOKCan_Click(Index As Integer)

    If Index = 0 Then
        RaiseEvent OK
      Else 'NOT INDEX...
        RaiseEvent Cancel
    End If

End Sub

Private Sub ChangedPosition(NewRow)

    If NewRow <> PreviousRow Then
        PreviousRow = NewRow
        RaiseEvent PositionChanged(NewRow)
    End If

End Sub

Public Property Get Cols() As Long
Attribute Cols.VB_Description = "Sets / returns the number of columns for the grid."
Attribute Cols.VB_HelpID = 10033

    Cols = gdBrowse.Cols - 1

End Property

Public Property Let Cols(ByVal nuCols As Long)

    If nuCols < 1 Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT NUCOLS...
        gdBrowse.Cols = nuCols + 1
        ReDim Preserve myColumnInfo(1 To nuCols)
        EqualColWidth
        gdBrowse.Col = 1
        gdBrowse.ColSel = gdBrowse.Cols - 1
    End If

End Property

Public Property Let ColWidth(ByVal Col As Long, ByVal nuColWidth As Long)
Attribute ColWidth.VB_Description = "Sets / returns the width in twips for a specific column."
Attribute ColWidth.VB_HelpID = 10038

  Dim i

    If Ambient.UserMode = False Then
        Err.Raise 387, Ambient.DisplayName
      Else 'NOT AMBIENT.USERMODE...
        If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
            Err.Raise 9, Ambient.DisplayName
          Else 'NOT COL...
            i = gdBrowse.ColWidth(gdBrowse.Cols - 1) + gdBrowse.ColWidth(Col) - nuColWidth
            If i < 120 Then
                i = 120
            End If
            myColumnInfo(gdBrowse.Cols - 1).myColWidth = i
            myColumnInfo(Col).myColWidth = nuColWidth
            SetColWidth
        End If
    End If

End Property

Public Property Get ColWidth(ByVal Col As Long) As Long

    If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT COL...
        ColWidth = myColumnInfo(Col).myColWidth
    End If

End Property

Public Property Get CurrentBookmark() As String

    CurrentBookmark = myCurrBookmark(gdBrowse.Row)

End Property

Public Property Get DisplayName() As String
Attribute DisplayName.VB_Description = "Sets / returns a user friendly name for the order-by field."
Attribute DisplayName.VB_HelpID = 10011

    DisplayName = lbSearchName

End Property

Public Property Let DisplayName(ByVal nuDisplayName As String)

    lbSearchName = nuDisplayName
    PropertyChanged "DisplayName"

End Property

Public Property Get DynamicScroll() As Boolean
Attribute DynamicScroll.VB_Description = "Sets / returns whether the srcollbar will dynamically scroll the grid."
Attribute DynamicScroll.VB_HelpID = 10006

    DynamicScroll = myDynamicScroll

End Property

Public Property Let DynamicScroll(ByVal nuDynamicScroll As Boolean)

    myDynamicScroll = (nuDynamicScroll <> False)
    PropertyChanged "DynamicScroll"

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Gibt einen Wert zurück, der bestimmt, ob ein Objekt auf vom Benutzer erzeugte Ereignisse reagieren kann, oder legt diesen fest."
Attribute Enabled.VB_HelpID = 10005

    Enabled = myEnabled

End Property

Public Property Let Enabled(ByVal nuEnabled As Boolean)

  Dim i

    myEnabled = nuEnabled
    If Not myRecordset Is Nothing Then
        For i = 0 To 5
            btNav(i).Enabled = nuEnabled
        Next i
        btOKCan(0).Enabled = nuEnabled
        btOKCan(1).Enabled = nuEnabled
        txSearch.Enabled = nuEnabled
        gdBrowse.Enabled = nuEnabled
    End If
    PropertyChanged "Enabled"

End Property

Public Sub EqualColWidth()

  Dim i, j, k

    With gdBrowse
        k = .Width - TpP - TpP
        .ColWidth(0) = DefWidth
        k = k - DefWidth
        j = Int(k / (.Cols - 1) / TpP) * TpP
        For i = 1 To .Cols - 2
            .ColWidth(i) = j
            myColumnInfo(i).myColWidth = j
            k = k - j
        Next i
        .ColWidth(i) = k - TpP - TpP - TpP - TpP
        myColumnInfo(i).myColWidth = k - TpP - TpP - TpP - TpP
    End With 'GDBROWSE

End Sub

Public Property Let FieldName(ByVal Col As Long, ByVal nuFieldName As String)
Attribute FieldName.VB_Description = "Sets / returns the fieldname for a specific column."
Attribute FieldName.VB_HelpID = 10018

    If Ambient.UserMode = False Then
        Err.Raise 387, Ambient.DisplayName
      Else 'NOT AMBIENT.USERMODE...
        If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
            Err.Raise 9, Ambient.DisplayName
          Else 'NOT COL...
            myColumnInfo(Col).myFieldName = nuFieldName
        End If
    End If

End Property

Public Property Get FieldName(ByVal Col As Long) As String

    If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT COL...
        FieldName = myColumnInfo(Col).myFieldName
    End If

End Property

Public Property Get FieldTranslation(ByVal Col As Long) As Boolean
Attribute FieldTranslation.VB_Description = "Turns field-translation for a specific column on or off."
Attribute FieldTranslation.VB_HelpID = 10019

    If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT COL...
        FieldTranslation = myColumnInfo(Col).myFieldTranslation
    End If

End Property

Public Property Let FieldTranslation(ByVal Col As Long, ByVal nuFieldTranslation As Boolean)

    If Ambient.UserMode = False Then
        Err.Raise 387, Ambient.DisplayName
      Else 'NOT AMBIENT.USERMODE...
        If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
            Err.Raise 9, Ambient.DisplayName
          Else 'NOT COL...
            myColumnInfo(Col).myFieldTranslation = (nuFieldTranslation <> False)
        End If
    End If

End Property

Private Sub FillGridFwd()

  Dim i

    gdBrowse.Clear
    gdBrowse.FormatString = Head
    SetColWidth
    With myRecordset
        NotFull = False
        For i = MinIx To MaxIx
            If .EOF Then
                .MoveLast
                lbFullPop.Visible = True
                NotFull = (.RecordCount >= MaxIx)
                Exit For '>---> Next
              Else '.EOF = FALSE
                myCurrBookmark(i) = .Bookmark
                gdBrowse.TextMatrix(i, 0) = .AbsolutePosition + 1
                FillRow i
                .MoveNext
            End If
        Next i
        FilledTo = i - 1
    End With 'MYRECORDSET
    If i > MaxIx Then
        SetScroll gdBrowse.TextMatrix(MaxIx, 0) - 1
    End If
    If myAutosize Then
        AdjustCols
    End If

End Sub

Private Sub FillGridRev()

  Dim i

    gdBrowse.Clear
    gdBrowse.FormatString = Head
    SetColWidth
    With myRecordset
        For i = MaxIx To MinIx Step -1
            If .BOF Then
                .MoveFirst
                Exit For '>---> Next
              Else '.BOF = FALSE
                gdBrowse.TextMatrix(i, 0) = .AbsolutePosition + 1
                myCurrBookmark(i) = .Bookmark
                FillRow i
                .MovePrevious
            End If
        Next i
    End With 'MYRECORDSET
    If i > 0 Then
        'that was not enough to fill the grid in reverse
        FillGridFwd
      Else 'NOT I...
        SetScroll gdBrowse.TextMatrix(MaxIx, 0) - 1
    End If
    If myAutosize Then
        AdjustCols
    End If

End Sub

Private Sub FillRow(RowNumber As Long)

  Dim i

    With myRecordset
        For i = 1 To gdBrowse.Cols - 1
            FieldContents = .Fields(myColumnInfo(i).myFieldName)
            If myColumnInfo(i).myFieldTranslation Then
                RaiseEvent TranslateColumn(myColumnInfo(i).myFieldName, FieldContents, FieldContents)
            End If
            If IsNull(FieldContents) Then
                FieldContents = "[?]"
            End If
            gdBrowse.TextMatrix(RowNumber, i) = Trim$(FieldContents)
        Next i
    End With 'MYRECORDSET

End Sub

Public Function FindFirst(Key As Variant) As Boolean

    With myRecordset
        If VarType(Key) = vbString Then
            .FindFirst myOrderedBy & " Like " & "'" & Key & "*'"
          Else 'NOT VARTYPE(KEY)...
            .FindFirst myOrderedBy & CompOper & Key
        End If
        If Not .NoMatch Then
            FillGridFwd
            FindFirst = True
            Hilite MinIx
        End If
    End With 'MYRECORDSET

End Function

Public Function FindLast(Key As Variant) As Boolean

    With myRecordset
        If VarType(Key) = vbString Then
            .FindLast myOrderedBy & " Like " & "'" & Key & "*'"
          Else 'NOT VARTYPE(KEY)...
            .FindLast myOrderedBy & OtherCompOper & Key
        End If
        If Not .NoMatch Then
            FillGridRev
            FindLast = True
            Hilite MaxIx
        End If
    End With 'MYRECORDSET

End Function

Public Property Let Forecolor(ByVal nuForecolor As OLE_COLOR)

    gdBrowse.Forecolor = nuForecolor
    PropertyChanged "ForeColor"

End Property

Public Property Get Forecolor() As OLE_COLOR

    Forecolor = gdBrowse.Forecolor

End Property

Public Property Get FullyPopulated() As Boolean

    FullyPopulated = lbFullPop.Visible

End Property

Private Sub gdBrowse_DblClick()

    btOKCan_Click 0

End Sub

Private Sub gdBrowse_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Hilite gdBrowse.Row

End Sub

Private Sub gdBrowse_SelChange()

    If gdBrowse.ColSel = gdBrowse.Cols - 1 Then
        ChangedPosition gdBrowse.Row
    End If

End Sub

Public Sub GoBookmark(Bookmark As String)

    myRecordset.Bookmark = Bookmark
    FillGridFwd

End Sub

Public Sub GoFirst()

    myRecordset.MoveFirst
    FillGridFwd
    With gdBrowse
        If .Row = 1 Then
            PreviousRow = 0
            ChangedPosition .Row
          Else 'NOT .ROW...
            Hilite MinIx
        End If
    End With 'GDBROWSE

End Sub

Public Sub GoLast()

    myRecordset.MoveLast
    lbFullPop.Visible = True
    FillGridRev
    With gdBrowse
        If .Row = MaxIx Then
            PreviousRow = 0
            ChangedPosition .Row
          Else 'NOT .ROW...
            Hilite FilledTo
        End If
    End With 'GDBROWSE

End Sub

Public Property Let HeadBackcolor(ByVal nuBackColor As OLE_COLOR)

    gdBrowse.BackColorFixed = nuBackColor
    PropertyChanged "HeadBackcolor"

End Property

Public Property Get HeadBackcolor() As OLE_COLOR

    HeadBackcolor = gdBrowse.BackColorFixed

End Property

Public Property Let Header(ByVal Col As Long, ByVal nuHeader As String)
Attribute Header.VB_Description = "Sets / returns the column header for a specific column."
Attribute Header.VB_HelpID = 10017

  Dim i

    If Ambient.UserMode = False Then
        Err.Raise 387, Ambient.DisplayName
      Else 'NOT AMBIENT.USERMODE...
        If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
            Err.Raise 9, Ambient.DisplayName
          Else 'NOT COL...
            myColumnInfo(Col).myHeader = nuHeader
            Head = ""
            For i = 1 To gdBrowse.Cols - 1
                Head = Head & "|" & myColumnInfo(i).myHeader
            Next i
            gdBrowse.FormatString = Head
            SetColWidth
        End If
    End If

End Property

Public Property Get Header(ByVal Col As Long) As String

    If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT COL...
        Header = myColumnInfo(Col).myHeader
    End If

End Property

Public Property Get HeadForecolor() As OLE_COLOR

    HeadForecolor = gdBrowse.ForeColorFixed

End Property

Public Property Let HeadForecolor(ByVal nuForecolor As OLE_COLOR)

    gdBrowse.ForeColorFixed = nuForecolor
    PropertyChanged "HeadForeColor"

End Property

Private Sub Hilite(Row As Long)

    With gdBrowse
        .Row = Row
        .Col = 1
        .ColSel = .Cols - 1
    End With 'GDBROWSE

End Sub

Public Property Let OrderedBy(ByVal nuOrderedBy As String)
Attribute OrderedBy.VB_Description = "Sets / returns the fieldname by which the recordset is ordered."
Attribute OrderedBy.VB_HelpID = 10015

    myOrderedBy = nuOrderedBy
    PropertyChanged "OrderedBy"

End Property

Public Property Get OrderedBy() As String

    OrderedBy = myOrderedBy

End Property

Public Property Get Recordset() As DAO.Recordset
Attribute Recordset.VB_Description = "Sets the recordset for the control; may be a bookmarkable dynaset or snapshot."
Attribute Recordset.VB_HelpID = 10014

    Set Recordset = myRecordset

End Property

Public Property Set Recordset(ByVal nuRecordset As DAO.Recordset)

    If Ambient.UserMode = False Then
        Err.Raise 383, Ambient.DisplayName
      Else 'NOT AMBIENT.USERMODE...
        If nuRecordset.Bookmarkable Then
            Set myRecordset = nuRecordset
            lbFullPop.Visible = False
            Enabled = myEnabled
            FillGridFwd
          Else 'NURECORDSET.BOOKMARKABLE = FALSE
            Err.Raise 300, Ambient.DisplayName, "Recordset is not bookmarkable"
        End If
    End If

End Property

Public Sub Refresh()

    myRecordset.Bookmark = myCurrBookmark(1)
    FillGridFwd

End Sub

Public Sub Reposition(Bookmark As String, ByVal Row As Long)

  Dim i

    gdBrowse.Visible = False
    GoBookmark Bookmark
    If Row < MinIx Or Row > MaxIx Then
        gdBrowse.Visible = True
        Err.Raise 380, Ambient.DisplayName, "Row number " & Row & " does not exist."
      Else 'NOT ROW...
        gdBrowse.Row = 1
        PageScroll = True
        Do Until Row <= 1
            Row = Row - 1
            ScrollLineRev
        Loop
        For i = MinIx To MaxIx
            If Bookmark = myCurrBookmark(i) Then
                gdBrowse.Row = i
                Exit For '>---> Next
            End If
        Next i
        gdBrowse.Col = 1
        gdBrowse.ColSel = gdBrowse.Cols - 1
        PageScroll = False
        gdBrowse.Visible = True
        If i > MaxIx Then
            Err.Raise 3159, Ambient.DisplayName, "Bookmark not found."
        End If
    End If

End Sub

Public Property Get RequiredColWidth(ByVal Col As Long) As Long

  Dim i, j, k

    If Col < LBound(myColumnInfo) Or Col > UBound(myColumnInfo) Then
        Err.Raise 9, Ambient.DisplayName
      Else 'NOT COL...
        With gdBrowse
            j = 0
            For i = MinIx - 1 To MaxIx
                k = UserControl.TextWidth(Trim$(.TextMatrix(i, Col))) + 120
                If k > j Then
                    j = k
                End If
            Next i
            RequiredColWidth = j
        End With 'GDBROWSE
    End If

End Property

Public Sub ScrollLineFwd()

  Dim i

    If gdBrowse.Row = FilledTo Then
        With myRecordset
            .Bookmark = myCurrBookmark(MaxIx)
            .MoveNext
            If .EOF Then
                .MoveLast
                lbFullPop.Visible = True
              Else '.EOF = FALSE
                gdBrowse.AddItem .AbsolutePosition + 1
                gdBrowse.RemoveItem 1
                For i = MinIx To MaxIx - 1
                    myCurrBookmark(i) = myCurrBookmark(i + 1)
                Next i
                myCurrBookmark(MaxIx) = .Bookmark
                FillRow i
                If Not PageScroll Then
                    PreviousRow = 0
                    ChangedPosition gdBrowse.Row
                    If myAutosize Then
                        AdjustCols
                    End If
                End If
            End If
            SetScroll Val(gdBrowse.TextMatrix(MaxIx, 0)) - 1
        End With 'MYRECORDSET
      Else 'NOT GDBROWSE.ROW...
        Hilite gdBrowse.Row + 1
    End If

End Sub

Public Sub ScrollLineRev()

  Dim i

    If gdBrowse.Row > 1 Then
        Hilite gdBrowse.Row - 1
      Else 'NOT GDBROWSE.ROW...
        With myRecordset
            .Bookmark = myCurrBookmark(1)
            .MovePrevious
            If .BOF Then
                .MoveFirst
              Else '.BOF = FALSE
                gdBrowse.AddItem .AbsolutePosition + 1, 1
                gdBrowse.RemoveItem MaxIx + 1
                For i = MaxIx To MinIx + 1 Step -1
                    myCurrBookmark(i) = myCurrBookmark(i - 1)
                Next i
                myCurrBookmark(1) = .Bookmark
                FillRow i
                If Not PageScroll Then
                    PreviousRow = 0
                    ChangedPosition gdBrowse.Row
                    If myAutosize Then
                        AdjustCols
                    End If
                End If
            End If
            SetScroll Val(gdBrowse.TextMatrix(MaxIx, 0)) - 1
        End With 'MYRECORDSET
    End If

End Sub

Public Sub ScrollPageFwd()

  Dim i

    With gdBrowse
        If .Row = MaxIx Then
            PreviousRow = 0
            ChangedPosition .Row
          Else 'NOT .ROW...
            Hilite FilledTo
        End If
        PageScroll = True
        For i = MinIx To MaxIx
            ScrollLineFwd
        Next i
        PageScroll = False
    End With 'GDBROWSE
    If myAutosize Then
        AdjustCols
    End If

End Sub

Public Sub ScrollPageRev()

  Dim i

    With gdBrowse
        If .Row = 1 Then
            PreviousRow = 0
            ChangedPosition .Row
          Else 'NOT .ROW...
            Hilite MinIx
        End If
        PageScroll = True
        For i = MinIx To MaxIx
            ScrollLineRev
        Next i
        PageScroll = False
    End With 'GDBROWSE
    If myAutosize Then
        AdjustCols
    End If

End Sub


Private Sub scScroll_Change()

    ScrChanged = True
    scScroll_Scroll
    ScrChanged = False

End Sub

Private Sub scScroll_Scroll()

    If myDynamicScroll Or ScrChanged Then
        If GetFocus = scScroll.hWnd Then
            myRecordset.Move scScroll * ScrollDivi - myRecordset.AbsolutePosition
            FillGridRev
            PreviousRow = 0
            ChangedPosition gdBrowse.Row
        End If
    End If

End Sub

Public Property Let SearchFor(nuSearch As String)
Attribute SearchFor.VB_Description = "Sets / returns the contents of the user accessible search key."
Attribute SearchFor.VB_HelpID = 10034

    txSearch = nuSearch

End Property

Public Property Get SearchFor() As String

    SearchFor = txSearch

End Property

Public Property Get SelectedRow() As Long

    SelectedRow = gdBrowse.Row

End Property

Private Sub SetColWidth()

  Dim i

    gdBrowse.ColWidth(0) = DefWidth
    For i = 1 To gdBrowse.Cols - 1
        gdBrowse.ColWidth(i) = myColumnInfo(i).myColWidth
    Next i

End Sub

Private Sub SetScroll(ByVal Value As Long)

    With scScroll
        Value = Value / ScrollDivi
        Do Until Value < 32767
            Value = Value / 2
            ScrollDivi = ScrollDivi * 2
            .Max = 0
        Loop
        If .Max < Value Then
            .Max = Value
            .LargeChange = Value / MaxIx
        End If
        If Value >= .Min And Value <= .Max Then
            If GetFocus <> scScroll.hWnd Then
                scScroll = Value
            End If
        End If
        .Enabled = (.Min < .Max)
    End With 'SCSCROLL

End Sub

Public Property Let SortOrder(ByVal nuSortOrder As SortDirection)
Attribute SortOrder.VB_Description = "Sets / returns the sort order of the SortedBy column."
Attribute SortOrder.VB_HelpID = 10016

    If nuSortOrder = Ascending Or nuSortOrder = Descending Then
        mySortOrder = nuSortOrder
        If mySortOrder = Ascending Then
            CompOper = " >= "
            OtherCompOper = " <= "
          Else 'NOT MYSORTORDER...
            CompOper = " <= "
            OtherCompOper = " >= "
        End If
        PropertyChanged "SortOrder"
      Else 'NOT NUSORTORDER...
        Err.Raise 380, Ambient.DisplayName
    End If

End Property

Public Property Get SortOrder() As SortDirection

    SortOrder = mySortOrder

End Property

Private Sub tmrTick_Timer()

    Select Case True
      Case ForwardLine
        ScrollLineFwd
      Case ReverseLine
        ScrollLineRev
      Case ForwardPage
        ScrollPageFwd
      Case ReversePage
        ScrollPageRev
    End Select
    tmrTick.Interval = 40

End Sub

Private Sub txSearch_Change()

  Dim BkMk   As String

    BkMk = myCurrBookmark(1)
    Select Case myRecordset.Fields(myOrderedBy).Type
      Case dbChar, dbText
        If Len(Trim$(txSearch)) = myRecordset.Fields(myOrderedBy).Size Then
            myRecordset.FindFirst myOrderedBy & CompOper & "'" & Trim$(txSearch) & "'"
          Else 'NOT LEN(TRIM$(TXSearch))...
            myRecordset.FindFirst myOrderedBy & " Like '" & Trim$(txSearch) & "*'"
        End If
      Case Else
        myRecordset.FindFirst myOrderedBy & CompOper & Trim$(txSearch)
    End Select
    If myRecordset.NoMatch Then
        Beep
    End If
    FillGridFwd
    If NotFull Then
        FillGridRev
    End If
    Hilite MinIx
    If myCurrBookmark(1) <> BkMk Then
        ChangedPosition gdBrowse.Row
    End If

End Sub

Private Sub txSearch_GotFocus()

    txSearch.Backcolor = &HC0FFFF

End Sub

Private Sub txSearch_LostFocus()

    txSearch.Backcolor = &HFFFFFF

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    lbSearchName.Backcolor = Ambient.Backcolor
    lbFullPop.Backcolor = Ambient.Backcolor
    If Ambient.UserMode = False Then
        gdBrowse.Text = Ambient.DisplayName
    End If

End Sub

Private Sub UserControl_Initialize()

    TpP = Screen.TwipsPerPixelX
    ScrollDivi = 1
    scScroll.Min = MaxIx - 1
    SetScroll MaxIx - 1

End Sub

Private Sub UserControl_InitProperties()

  Dim i

    myEnabled = True
    mySortOrder = Ascending
    CompOper = " >= "
    OtherCompOper = " <= "
    gdBrowse.Cols = 10
    ReDim myColumnInfo(1 To gdBrowse.Cols - 1)
    For i = 1 To gdBrowse.Cols - 1
        myColumnInfo(i).myColWidth = 960
        myColumnInfo(i).myHeader = ""
    Next i
    gdBrowse.BackColorSel = vbHighlight
    gdBrowse.ForeColorSel = vbHighlightText
    gdBrowse.Backcolor = vbWindowBackground
    gdBrowse.Forecolor = vbWindowText
    gdBrowse.BackColorFixed = vbButtonFace
    gdBrowse.ForeColorFixed = vbButtonText
    EqualColWidth

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

  Dim i       As Integer
  Dim Focus   As Long
  Static Busy As Boolean

    i = InStr(NavCodes, Chr$(KeyCode)) - 1
    If i >= 0 And i <= 5 Then
        KeyCode = 0
        If GetFocus <> gdBrowse.hWnd And Not Busy Then
            'the grid consumes keydown before preview - no chance to catch it
            If btNav(i).Enabled Then
                Busy = True
                Focus = GetFocus
                btNav_MouseDown i, 0, 0, 0, 0
                btNav_MouseUp i, 0, 0, 0, 0
                PutFocus Focus
                DoEvents
                Busy = False
            End If
        End If
    End If

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    If Chr$(KeyAscii) = " " And GetFocus <> txSearch.hWnd Then
        btOKCan_Click 0
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        myEnabled = .ReadProperty("Enabled", True)
        myAutosize = .ReadProperty("Autosize", False)
        myDynamicScroll = .ReadProperty("DynamicScroll", False)
        myDisplayName = .ReadProperty("DisplayName", "")
        myOrderedBy = .ReadProperty("OrderedBy", "")
        SortOrder = .ReadProperty("SortOrder", Ascending)
        Cols = .ReadProperty("Cols", 10)
        gdBrowse.BackColorSel = .ReadProperty("BarBackcolor", vbHighlight)
        gdBrowse.ForeColorSel = .ReadProperty("BarForecolor", vbHighlightText)
        gdBrowse.Backcolor = .ReadProperty("Backcolor", vbWindowBackground)
        gdBrowse.Forecolor = .ReadProperty("Forecolor", vbWindowText)
        gdBrowse.BackColorFixed = .ReadProperty("HeadBackcolor", vbButtonFace)
        gdBrowse.ForeColorFixed = .ReadProperty("HeadForecolor", vbButtonText)
    End With 'PROPBAG
    EqualColWidth

End Sub

Private Sub UserControl_Resize()

    ReDim myColumnInfo(1 To gdBrowse.Cols - 1)
    Size UserControl.Width, btNav(0).Top + btNav(0).Height
    scScroll.Left = UserControl.Width - scScroll.Width
    gdBrowse.Width = scScroll.Left
    EqualColWidth

End Sub

Private Sub UserControl_Show()

  Dim i

    UserControl_AmbientChanged ""
    For i = MinIx To MaxIx
        gdBrowse.TextMatrix(i, 0) = i
    Next i
    If Ambient.UserMode = False Then
        gdBrowse.Text = Ambient.DisplayName
    End If

End Sub

Private Sub UserControl_Terminate()

    If Not myRecordset Is Nothing Then
        myRecordset.Close
        Set myRecordset = Nothing
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        .WriteProperty "Enabled", myEnabled, True
        .WriteProperty "Autosize", myAutosize, False
        .WriteProperty "DynamicScroll", myDynamicScroll, False
        .WriteProperty "DisplayName", myDisplayName, ""
        .WriteProperty "OrderedBy", myOrderedBy, ""
        .WriteProperty "SortOrder", mySortOrder, Ascending
        .WriteProperty "Cols", gdBrowse.Cols - 1, 10
        .WriteProperty "BarBackcolor", gdBrowse.BackColorSel, vbHighlight
        .WriteProperty "BarForecolor", gdBrowse.ForeColorSel, vbHighlightText
        .WriteProperty "Backcolor", gdBrowse.Backcolor, vbWindowBackground
        .WriteProperty "Forecolor", gdBrowse.Forecolor, vbWindowText
        .WriteProperty "HeadBackcolor", gdBrowse.BackColorFixed, vbButtonFace
        .WriteProperty "HeadForecolor", gdBrowse.ForeColorFixed, vbButtonText
    End With 'PROPBAG

End Sub

':) Ulli's VB Code Formatter V2.11.3 (09.04.2002 17:47:50) 54 + 1093 = 1147 Lines
