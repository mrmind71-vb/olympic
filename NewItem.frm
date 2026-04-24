VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form NewItem 
   Caption         =   "ÊÓÌíá ÕäÝ ÌÏíÏ"
   ClientHeight    =   4440
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar BAR 
      Height          =   195
      Left            =   675
      TabIndex        =   24
      Top             =   4185
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "fix code"
      Height          =   465
      Left            =   2565
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3690
      Width           =   1725
   End
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÍÝÙ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5400
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3690
      UseMaskColor    =   -1  'True
      Width           =   1410
   End
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   675
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3690
      UseMaskColor    =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Height          =   3480
      Left            =   630
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   90
      Width           =   6135
      Begin VB.CheckBox xrep 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊßÑÇÑ ÇáÕäÝ"
         Height          =   285
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3060
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.TextBox XITEM2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   21
         Top             =   180
         Width           =   1335
      End
      Begin VB.TextBox XPRICE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   525
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2610
         Width           =   1230
      End
      Begin VB.TextBox XCOST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3105
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2610
         Width           =   1230
      End
      Begin VB.TextBox XBAR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3105
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "2"
         Top             =   2205
         Width           =   1230
      End
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   585
         Width           =   4200
      End
      Begin VB.TextBox xItem 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1875
         MaxLength       =   15
         TabIndex        =   9
         Top             =   180
         Width           =   2460
      End
      Begin VB.TextBox xRATE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   525
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Text            =   "30"
         Top             =   2190
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   360
         Left            =   540
         TabIndex        =   1
         Top             =   990
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   360
         Left            =   540
         TabIndex        =   2
         Top             =   1395
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xSupler 
         Height          =   360
         Left            =   540
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1800
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÈíÚ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2670
         Width           =   840
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÚÑ ÇáÊßáÝÉ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2670
         Width           =   1125
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãæÑÏ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1890
         Width           =   525
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "äæÚ ØÈÇÚÉ ÈÇÑßæÏ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2250
         Width           =   1290
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãÌãæÚÉ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÑÞã ÇáÕäÝ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÞÓã "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1080
         Width           =   570
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÅÓã ÇáÕäÝ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   675
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "äÓÈÉ ãÓÊåáß"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2250
         Width           =   1125
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   585
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
End
Attribute VB_Name = "NewItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sItem As String
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    If Len(xItem.Text) > 6 Or Not IsNumeric(xItem.Text) Then XITEM2.Text = ""
    If Len(xItem.Text) <= 6 And XITEM2.Text = "" And IsNumeric(xItem.Text) Then
        If MsgBox("ÊÔÛíá ÇáÊÑÞíã ÇáÊáÞÇÆì", vbYesNo) = vbYes Then XITEM2.Text = xItem.Text
    End If
    If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(xItem.Text)) = "" Then
        cStr1 = "insert into FILE1_10(ITEM,ITEM2,DESCA,[GROUP],[SECTION],[SUPLER] ,  [RATE] , T_BAR ,COST,PRICE)" & _
        "VALUES(" & _
        addstring(xItem.Text) & "," & _
        addvalue(XITEM2.Text) & "," & _
        addstring(xDescA.Text) & "," & _
        addvalue(xgroup.BoundText) & "," & _
        addvalue(xSection.BoundText) & "," & _
        addstring(xSupler.BoundText) & "," & _
        addvalue(xRATE.Text) & "," & _
        addvalue(XBAR.Text) & "," & _
        Val(XCOST.Text) & "," & _
        Val(XPRICE.Text) & _
        ")"
        con.Execute cStr1
    End If
    Unload Me
End Sub
Private Sub Command1_Click()
    Dim datatable As New ADODB.Recordset
    Dim nRec As Double
    If MsgBox("ÓæÝ íÊã ÖÈØ ÇáÊßæíÏ ÇáÊáÞÇÆì", vbOKCancel) = vbOK Then
    BAR.Visible = True
    BAR.Min = 0
    BAR.Value = 0
    con.BeginTrans
    con.Execute " update file1_10 set item2 = null "
    con.CommitTrans
    datatable.Open "FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdTable
    BAR.Max = datatable.RecordCount
    With datatable
        .MoveFirst
        Do While Not .EOF
            nRec = nRec + 1
            BAR.Value = nRec
            If IsNumeric(!Item) And Len(!Item) <= 6 Then
                con.Execute " update file1_10 set item2 = ITEM WHERE ITEM = " & MyParn(!Item)
            End If
            .MoveNext
        Loop
    End With
    MsgBox "Êã ÇáÖÈØ"
    End If
End Sub
Private Sub Form_Load()
openCon con

DATA1.ConnectionString = strCon
DATA1.RecordSource = "FILE1_50"
Set xgroup.RowSource = DATA1
xgroup.ListField = "Desca"
xgroup.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE1_10SC"
Set xSection.RowSource = DATA2
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "FILE4_10"
Set xSupler.RowSource = DATA3
xSupler.ListField = "Desca"
xSupler.BoundColumn = "Code"

xDescA.Text = ""
xgroup.BoundText = ""
xSection.BoundText = ""
xSupler.BoundText = ""
xRATE.Text = 30
XBAR.Text = 1
XCOST.Text = ""
XPRICE.Text = ""

If sItem <> "" Then
    Dim aret As Variant
    aret = aGetDesca("select DESCA,[GROUP],[SECTION],[SUPLER] ,  [RATE] , T_BAR ,COST,PRICE  from file1_10 where item = " & MyParn(sItem))
    xItem.Text = sItem
    If UBound(aret) <> 0 Then
        xDescA.Text = aret(1) & ""
        xgroup.BoundText = aret(2) & ""
        xSection.BoundText = aret(3) & ""
        xSupler.BoundText = aret(4) & ""
        xRATE.Text = aret(5) & ""
        XBAR.Text = aret(6) & ""
        XCOST.Text = aret(7) & ""
        XPRICE.Text = aret(8) & ""
    End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    sItem = ""
    closeCon con
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
If KeyAscii = 19 Then
    cmdSave_Click
End If
End Sub
Private Sub XCOST_Change()
    XPRICE.Text = myNear(Round(Val(XCOST.Text) * ((100 + Val(xRATE.Text)) / 100), 2), 0.5)
End Sub
Private Sub xitem_LostFocus()
    If Len(xItem.Text) <= 7 Then
    If MsgBox("ÊÍÏíË ÇáÊÑÞíã ÇáÊáÞÇÆì ÈåÐÇ ÇáßæÏ", vbYesNo) = vbYes Then
        Purchasefrm.grid1.TextMatrix(Purchasefrm.grid1.Row, 1) = xItem.Text
        XITEM2.Text = xItem.Text
        con.Execute " UPDATE FILE1_10 SET ITEM2 = " & addvalue(XITEM2.Text) & " WHERE ITEM = " & MyParn(xItem.Text)
    End If
    End If
End Sub
