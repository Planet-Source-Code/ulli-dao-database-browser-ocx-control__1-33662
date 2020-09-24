VERSION 5.00
Object = "{EBB8A42E-324D-11D4-B07A-8B3DAE15DB09}#32.0#0"; "Browser.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Browser Example"
   ClientHeight    =   6285
   ClientLeft      =   705
   ClientTop       =   2010
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton Command4 
      Caption         =   "Reset col width"
      Height          =   390
      Left            =   645
      TabIndex        =   5
      Top             =   5670
      Width           =   1680
   End
   Begin BrowserOCX.Browser Browser1 
      Height          =   4290
      Left            =   135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   75
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   7567
      Cols            =   9
      BarBackcolor    =   192
      BarForecolor    =   8454143
      Backcolor       =   12640511
      HeadBackcolor   =   8421376
      HeadForecolor   =   16777152
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Set opt col width"
      Height          =   390
      Left            =   645
      TabIndex        =   3
      Top             =   5190
      Width           =   1680
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reposition"
      Enabled         =   0   'False
      Height          =   390
      Left            =   6660
      TabIndex        =   2
      Top             =   5175
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Set equal col width "
      Height          =   390
      Left            =   645
      TabIndex        =   1
      Top             =   4710
      Width           =   1680
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Item"
      Height          =   195
      Left            =   3225
      TabIndex        =   6
      Top             =   4875
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   3
      X1              =   75
      X2              =   10125
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3210
      TabIndex        =   0
      Top             =   5175
      Width           =   3180
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ws              As Workspace
Private db              As Database
Private rs              As Recordset
Private PrevBookmark    As String
Private PrevRow         As Long

Private Sub Command1_Click()

  'equal columns widths

    Browser1.EqualColWidth

End Sub

Private Sub Command2_Click()

  'reposition browser to a known bookmark

    Browser1.Reposition PrevBookmark, PrevRow

End Sub

Private Sub Command3_Click()

  'adjust columns to fit contents

    Browser1.AdjustCols

End Sub

Private Sub Command4_Click()

  'individual column widths

    With Browser1
        .ColWidth(1) = 600
        .ColWidth(2) = 2040
        .ColWidth(3) = 2700
        .ColWidth(4) = 1800
        .ColWidth(5) = 1500
        .ColWidth(6) = 540
    End With 'BROWSER1

End Sub

Private Sub Form_Load()

    Set ws = CreateWorkspace("", "Admin", "", dbUseJet)
    Set db = ws.OpenDatabase("c:\programme\microsoft visual studio\vb98\biblio.mdb") 'or wherever the database is on your computer
    Set rs = db.OpenRecordset("SELECT * FROM Publishers ORDER BY Name", dbOpenSnapshot)

    With Browser1
    
        'customize browser
        .Cols = 6
        .Header(1) = ">Id" 'the >prefix aligns the text right
        .Header(2) = "Name"
        .Header(3) = "Company"
        .Header(4) = "Address"
        .Header(5) = "City"
        .Header(6) = "State"

        .FieldName(1) = "PubID"
        .FieldName(2) = "Name"
        .FieldName(3) = "Company Name"
        .FieldName(4) = "Address"
        .FieldName(5) = "City"
        .FieldName(6) = "State"

        .FieldTranslation(4) = True
        .FieldTranslation(5) = True

        .OrderedBy = .FieldName(2)
        .SortOrder = Descending
        .DisplayName = "Comp.Name"
        .DynamicScroll = True 'follow scrollbar (for faster machines)

        'fill browser
        Set .Recordset = rs.Clone

        Command4_Click 'sets individual columns widths

    End With 'BROWSER1

End Sub

Private Sub Browser1_TranslateColumn(ByVal FieldName As String, ByVal OldValue As Variant, NewValue As Variant)

  'column translation

    Select Case FieldName
      Case "City", "Address"
        If IsNull(OldValue) Then
            NewValue = "/unknown/"
        End If
    End Select

End Sub

Private Sub Browser1_OK()

  'user clicked browser OK button

    rs.Bookmark = Browser1.CurrentBookmark
    Label1 = rs.Fields("PubID") & " " & rs.Fields("Name")
    PrevBookmark = Browser1.CurrentBookmark
    PrevRow = Browser1.SelectedRow
    Command2.Enabled = True

End Sub

Private Sub Browser1_Cancel()

  'user clicked browser Cancel button

    Label1 = ""
    Command2.Enabled = False

End Sub

':) Ulli's VB Code Formatter V2.11.3 (10.04.2002 09:12:24) 6 + 115 = 121 Lines
