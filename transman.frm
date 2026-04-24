VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form transManfrm 
   Caption         =   "ÊÍæíáÇÊ"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2805
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkprint 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáÛÇÁ ÇáØÈÇÚÉ"
      Height          =   195
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2565
      Width           =   1815
   End
   Begin VB.TextBox xcode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   5355
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2115
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "ÍÝÙ"
      Height          =   420
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2115
      Width           =   1545
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "ÎÑæÌ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2115
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Height          =   2040
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   6675
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   630
         Left            =   135
         MaxLength       =   200
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1305
         Width           =   5520
      End
      Begin VB.TextBox xvalue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4230
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo xBox1 
         Height          =   315
         Left            =   2925
         TabIndex        =   0
         Top             =   225
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xBox2 
         Height          =   315
         Left            =   2925
         TabIndex        =   1
         Top             =   585
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "ÇáÈíÇä :"
         Height          =   330
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1395
         Width           =   690
      End
      Begin VB.Label Label3 
         Caption         =   "ÇáãÈáÛ :"
         Height          =   330
         Left            =   5805
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Åáí ÎÒäÉ :"
         Height          =   330
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "ãä ÎÒäÉ :"
         Height          =   330
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   870
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   7500
      Left            =   675
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   5385
      _cx             =   9499
      _cy             =   13229
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   0   'False
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   1440
      MarginTop       =   1440
      MarginRight     =   1440
      MarginBottom    =   1440
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   40.6801007556675
      ZoomMode        =   4
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
End
Attribute VB_Name = "transManfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sBox1 As String, Sbox2 As String, sCaption As String, sDate As String
Private Sub chkprint_Click()
addSetting "print_trans", chkprint.Value, App.Path & App.Path & "\other.txt"
End Sub
Private Sub cmdExit_Click()
Unload transManfrm
End Sub
Private Sub cmdSave_Click()
mysave
End Sub
Private Sub Form_Load()
chkprint.Value = Val(RetSetting("print_trans", App.Path & "\other.txt"))
openCon con
cString = "Select file0_50.* From file0_50"
DATA1.ConnectionString = strCon
DATA1.RecordSource = cString

Set xBox1.RowSource = DATA1
xBox1.ListField = "Desca"
xBox1.BoundColumn = "Code"

Set xBox2.RowSource = DATA1
xBox2.ListField = "Desca"
xBox2.BoundColumn = "Code"
myload
End Sub
Private Sub myload()
xBox1.BoundText = sBox1
xBox2.BoundText = Sbox2
xBox1_LostFocus
xBox2_LostFocus
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set transManfrm = Nothing
closeCon con
End Sub

Private Sub xBox1_LostFocus()
If Not xBox1.MatchedWithList Then xBox1.BoundText = ""
End Sub
Private Sub xBox2_LostFocus()
If Not xBox2.MatchedWithList Then xBox2.BoundText = ""
End Sub
Private Function MYVALID() As Boolean
If xBox1.BoundText = "" Then
    MsgBox "áÇ ÊæÌÏ ÎÒíäÉ Çæáí"
    Exit Function
End If
If xBox2.BoundText = "" Then
    MsgBox "áÇ ÊæÌÏ ÎÒíäÉ ËÇäíÉ"
    Exit Function
End If
If Val(xvalue.Text) = 0 Then
    MsgBox "áÇ ÊæÌÏ ÞíãÉ áÇÖÇÝÊåÇ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub mysave()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform "Êã ÍÝÙ ÇáÈíÇäÇÊ ÈäÌÇÍ"
vp.Visible = True
If Me.chkprint.Value = 0 Then
    For i = 1 To 2
        doprint
    Next
End If
End Sub
Private Function myreplace() As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xcode.Text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xdesca.Text)

aInsert(2, 0) = "Date"
aInsert(2, 1) = addDate(sDate)

aInsert(3, 0) = "NO1"
aInsert(3, 1) = addstring(xBox1.BoundText)

aInsert(4, 0) = "NO2"
aInsert(4, 1) = addstring(xBox2.BoundText)

aInsert(5, 0) = "[VALUE]"
aInsert(5, 1) = Val(xvalue.Text)


On Error GoTo myerror
con.BeginTrans
xcode.Text = RetZero(Val(Newflag("FILE0_51", "CODE")), 3)
aInsert(0, 1) = addstring(xcode.Text)
con.Execute CreateInsert(aInsert, "FILE0_51")
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub doprint()
With vp
     vp = " "
    .Device = pDevice
    .StartDoc
    .fontsize = 8
    .TextColor = vbBlack
    .FontName = "Arial"
    .MarginLeft = 150
    .TextAlign = taCenterTop
    .PenStyle = psTransparent
    
    .Paragraph = "ÊÍæíá ãä ÎÒäÉ : " & xBox1.Text
    .Paragraph = "Çáí  ÎÒäÉ : " & xBox2.Text
    .TextAlign = taRightTop
    .Paragraph = String(40, "=")
    .Paragraph = "ÊÇÑíÎ : " & Format(Date, "DD-MM-YYYY")
    .Paragraph = "æÞÊ : " & Time
    .Paragraph = String(40, "=")
    
    f = ">3000|<600;"
    H = ""
    
    cRow = xvalue.Text & "|" & _
    ": ÇáãÈáÛ" & ";"
    .AddTable f, H, cRow
    
    cRow = xdesca.Text & "|" & _
          ": ÇáÈíÇä" & ";"
    .AddTable f, H, cRow
    
    .EndDoc
End With
End Sub

