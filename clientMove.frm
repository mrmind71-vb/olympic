VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form clientMoveFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·⁄„·«¡"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   900
      Top             =   585
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
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
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   270
      Width           =   4920
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "clientMove.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "clientMove.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "clientMove.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "clientMove.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   8190
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   6675
      Begin VB.TextBox XDATE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2025
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Tag             =   "D"
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Tag             =   "D"
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XCODE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   945
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   3300
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7575
      Left            =   90
      TabIndex        =   8
      Top             =   1035
      Width           =   14775
      _cx             =   26061
      _cy             =   13361
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame5 
      Height          =   1275
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8595
      Width           =   14775
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   1005
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Width           =   14595
         _cx             =   25744
         _cy             =   1773
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   16
      Top             =   9900
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "ClientMoveFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ClientTable As New ADODB.Recordset
Public Sub myloadgrd()
Dim cString As String, nPrevious As Double, nFirst_Bal As Double
Dim loctable As New ADODB.Recordset
If IsDate(xdate1.Text) Then
   cString = "Select sum([SAL] - PAY)  as Balance from FILE3_11 where FILE3_11.CODE = " & MyParn(xCode.Text) & _
              " and FILE3_11.date < " & DateSq(xdate1.Text)
   nFirst_Bal = Round(Val(GetField(cString, con) & ""), 2)
   nPrevious = nFirst_Bal
End If

cString = "select FILE3_11.*,file3_12.desca as moveDesca  " & _
          " From FILE3_11 Left join file3_12 on FILE3_11.type = file3_12.code"

cString = cString & turn(cString) & " FILE3_11.code = " & MyParn(xCode.Text)

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " FILE3_11.date >= " & DateSq(xdate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & " FILE3_11.date <= " & DateSq(xDate2.Text)
End If
cString = cString & " Order by FILE3_11.date,file3_12.[order],FILE3_11.doc_id,pay"
With grid1
    .Rows = 1
   If nPrevious <> 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
        .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "Fixed")
   End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Round(Val(loctable!sal & ""), 2) - Round(Val(loctable!Pay & ""), 2)
        .TextMatrix(.Rows - 1, 0) = loctable!MoveDesca
        If Trim(loctable!desca) & "" <> "" Then
            .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 1, 0) & turn(.TextMatrix(.Rows - 1, 0), "-") & loctable!desca
        End If
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!doc_ID & ""
        .TextMatrix(.Rows - 1, 3) = Myvalue(loctable!sal, "fixed")
        .TextMatrix(.Rows - 1, 4) = Myvalue(loctable!Pay, "fixed")
        .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
        .TextMatrix(.Rows - 1, 6) = loctable!Type & ""
        loctable.MoveNext
    Loop
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
    If grid1.Rows > 1 Then
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
    End If
End With
StatusBar1.Panels(1).Text = "—’Ìœ «·⁄„Ì· : " & Round(nPrevious, 2)
loctable.Close
Set loctable = Nothing

cField = myiif("TYPE = '1'", "SAL-PAY") & " AS FIRST_BAL"
cField = cField & "," & _
        myiif("TYPE = '4'", "SAL") & " AS SALES"

cField = cField & "," & _
        myiif("TYPE = '7'", "PAY") & " AS PAY_FROM"

cField = cField & "," & _
        myiif("TYPE = '8'", "SAL") & " AS PAY_TO"

cField = cField & "," & _
        myiif("TYPE = 'A'", "PAY") & " AS [CHECK]"

cField = cField & "," & _
        myiif("TYPE = 'C'", "SAL") & " AS CHECK_REF"

cField = cField & "," & _
        myiif("TYPE = 'E'", "SAL") & " AS CHECK_TRANS"

cString = "SELECT " & cField & _
          " FROM FILE3_11"
cString = cString & turn(cString) & " CODE = " & MyParn(xCode.Text)
If IsDate(xdate1.Text) Then cString = cString & turn(cString) & " DATE >= " & addDate(xdate1.Text)
If IsDate(xDate2.Text) Then cString = cString & turn(cString) & " DATE <= " & addDate(xDate2.Text)

Dim aRet As Variant
aRet = GetFields(cString)
If IsEmpty(aRet) Then Exit Sub
'.TextMatrix(2, 4) = "√Ê—«ﬁ ﬁ»÷ €Ì— „Õ’·…"

With grid2
    .TextMatrix(0, 1) = Myvalue(nFirst_Bal + Val(retFlag(aRet, "FIRST_BAL") & ""), "FIXED")
    .TextMatrix(0, 3) = Myvalue(retFlag(aRet, "SALES"), "FIXED")
    .TextMatrix(0, 5) = Myvalue(Val(retFlag(aRet, "PAY_TO") & ""), "FIXED")
    .TextMatrix(0, 7) = Myvalue(Val(retFlag(aRet, "CHECK_REF") & "") + Val(retFlag(aRet, "CHECK_TRANS") & ""), "FIXED")
    .TextMatrix(0, 9) = Myvalue(Val(.TextMatrix(0, 1)) + Val(.TextMatrix(0, 3)) + Val(.TextMatrix(0, 5)) + Val(.TextMatrix(0, 7)), "FIXED")
    
    .TextMatrix(1, 1) = Myvalue(Val(retFlag(aRet, "PAY_FROM") & ""), "FIXED")
    .TextMatrix(1, 3) = Myvalue(Val(retFlag(aRet, "CHECK") & ""), "FIXED")
    .TextMatrix(1, 9) = Myvalue(Val(.TextMatrix(1, 1)) + Val(.TextMatrix(1, 3)), "FIXED")
    .TextMatrix(2, 9) = Myvalue(Val(.TextMatrix(0, 9)) - Val(.TextMatrix(1, 9)), "FIXED")
End With
End Sub
Public Sub myloadgrd2()
Dim cStrW As String
Dim loctable As New ADODB.Recordset
Dim datatable As New ADODB.Recordset
Dim n11 As Double
Dim n12 As Double
Dim n13 As Double
Dim n14 As Double
Dim n15 As Double
Dim n16 As Double
Dim n17 As Double
cString = "select FILE3_11.DATE,MIN(DOC_ID) + ' - ' + MAX(DOC_ID) as MINMAXID,SUM(PAY) AS SUMOFPAY,SUM(SAL) AS SUMOFSAL,file3_12.desca,FILE3_11.TYPE " & _
          " From FILE3_11 Left join file3_12 on FILE3_11.type = file3_12.code"

cString = cString & turn(cString) & " FILE3_11.code = " & MyParn(xCode.Text)

If IsDate(xdate1.Text) Then
    cString = cString & turn(cString) & " FILE3_11.date >= " & DateSq(xdate1.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turn(cString) & " FILE3_11.date <= " & DateSq(xDate2.Text)
End If

cString = cString & " Group by FILE3_11.date,file3_11.[type],file3_12.desca"
cString = cString & " Order by FILE3_11.date,min(file3_12.[order]),min(FILE3_11.doc_id),min(pay)"
With grid1
    .Rows = 1
    If IsDate(xdate1.Text) Then
       cString2 = "Select sum([SAL] - PAY)  as Balance from FILE3_11RP where FILE3_11RP.CODE = " & MyParn(xCode.Text) & _
                  " and FILE3_11RP.date < " & DateSq(xdate1.Text)
       nPrevious = Round(Val(GetDesca(cString2)))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.Text
            .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "Fixed")
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         grid1.AddItem ""
         nPrevious = nPrevious + Round(Val(loctable!sumofsal & ""), 2) - Round(Val(loctable!SumofPay & ""), 2)
        .TextMatrix(.Rows - 1, 0) = loctable!desca & ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!MINMAXID & ""
        .TextMatrix(.Rows - 1, 3) = TurnValue(Round(Val(loctable!sumofsal & ""), 2), 0, "")
        .TextMatrix(.Rows - 1, 4) = TurnValue(Round(Val(loctable!SumofPay & ""), 2), 0, "")
        .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
        .TextMatrix(.Rows - 1, 6) = loctable!Type
        loctable.MoveNext
    Loop
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
    .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
    .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
End With
StatusBar1.Panels(1).Text = Round(nPrevious, 2)

If IsDate(xdate1.Text) Then
    cStrW = cStrW & " AND DATE >= " & DateSq(xdate1.Text)
End If
If IsDate(xDate2.Text) Then
    cStrW = cStrW & " AND DATE <= " & DateSq(xDate2.Text)
End If


cStr1 = " SELECT SUM(SAL + PAY )  AS VALMOVE , [TYPE] FROM FILE3_11 WHERE CODE = " & MyParn(xCode.Text) & cStrW & " group by [type] "
datatable.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText

n11 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE3_11 WHERE [TYPE] = '4' AND CODE = " & MyParn(xCode.Text) & cStrW) & "")
n12 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '5' AND CODE = " & MyParn(xCode.Text) & cStrW) & "")
n13 = n11 - n12
n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = '10' OR [TYPE] = '11' OR [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(xCode.Text) & cStrW) & "")



n13 = n11 - n12
n16 = n14 + n15
n17 = Val(GetDesca("SELECT SUM(VALUE)  FROM FILE5_20 WHERE [CLOSED] = '0' AND CODE1 = " & MyParn(xCode.Text)) & "")
With grid2
    .Rows = 3
    .Cols = 6
    .FixedCols = 0
    .FixedRows = 0
    .TextMatrix(0, 0) = "Ã. „»Ì⁄« "
    .TextMatrix(0, 2) = "Ã. „— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï „»Ì⁄«  "

    .TextMatrix(1, 0) = "œ›⁄«  ‰ﬁœÏ"
    .TextMatrix(1, 2) = "œ›⁄«  ‘Ìﬂ« "
    .TextMatrix(1, 4) = "≈Ã„«·Ï œ›⁄« "
    .TextMatrix(2, 4) = "√Ê—«ﬁ ﬁ»÷ €Ì— „Õ’·…"
    
    .TextMatrix(0, 1) = n11
    .TextMatrix(0, 3) = n12
    .TextMatrix(0, 5) = n13

    .TextMatrix(1, 1) = n14
    .TextMatrix(1, 3) = n15
    .TextMatrix(1, 5) = n16
    
    .TextMatrix(2, 5) = n17
    For i = 0 To 5
        .ColWidth(i) = 1500
    Next i
End With
End Sub

Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
End Sub
Function MYVALID() As Boolean
If Trim(xCode.Text) = "" Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— „”Ã·"
    Exit Function
End If
'If GetDesca("select Desca from file3_10 where code = " & MyParn(xCode.Text)) = "" Then
'    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
'    Exit Function
'End If
If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub

Private Sub Check1_Click()
If Not MYVALID Then Exit Sub
myloadgrd
End Sub

Private Sub cmdExel_Click()
ToFileExel grid1
End Sub

Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
myloadgrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdPrint_Click()
doprint
End Sub
Private Sub Form_Load()
With grid1
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "„œÌ‰"
.TextMatrix(0, 4) = "œ«∆‰"
.TextMatrix(0, 5) = "—’Ìœ"

.ColWidth(0) = 5000
.ColWidth(1) = 1500
.ColWidth(2) = 2000
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
.ColHidden(6) = True
End With
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
fixgrd2
openCon con
'On Error Resume Next
'xCode.SetFocus
'Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
On Error Resume Next
Unload Search3
Err.Clear
End Sub

Private Sub Grid1_DblClick()
    Dim cDoc_no As String
    Select Case grid1.TextMatrix(grid1.Row, 6)
        Case "4", "5"
            cDoc_no = grid1.TextMatrix(grid1.Row, 2)
            salesfrm.myPublic = IIf(grid1.TextMatrix(grid1.Row, 6) = "4", 0, 1)
            salesfrm.sDoc_no = cDoc_no
'            SalesFrm.Frame9.Visible = False
'            SalesFrm.Frame1.Visible = False
            salesfrm.Show
        Case "A", "C"
            chqClientfrm.sSer_no = grid1.TextMatrix(grid1.Row, 2)
            chqClientfrm.Show 1
    End Select
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub

Private Sub xCode_Change()
grid1.Rows = 1
cmdGo.Enabled = Trim(xCode.Text) <> ""
End Sub

Private Sub XCODE_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub

Private Sub xCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xDesca.Caption = ""
If Trim(xCode.Text) = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
xDesca.Caption = GetDesca("select Desca from file3_10 where code = " & MyParn(xCode.Text))
End Sub
Private Sub xStore_Click(Area As Integer)
If Not cmdGo.Enabled Then cmdGo.Enabled = True
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file3_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub doprint()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim n11 As Double, n12 As Double, n13 As Double, n14 As Double, n15 As Double, n16 As Double, n17 As Double
Dim cStrW As String

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If Trim(xCode.Text) <> "" Then
    aHeader(0) = "[" & "··⁄„Ì· : " & xDesca.Caption & "]"
End If
If IsDate(xdate1.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
If IsDate(xDate2.Text) Then
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If
With grid1
For i = 1 To .Rows - 2
    temptable.AddNew
    temptable!Date1 = TurnValue(.TextMatrix(i, 1), "yyyy-mm-dd")
    temptable!str1 = TurnValue(.TextMatrix(i, 2))
    temptable!str2 = TurnValue(.TextMatrix(i, 0))
    temptable!val1 = Val(.TextMatrix(i, 3))
    temptable!val2 = Val(.TextMatrix(i, 4))
    temptable!Val3 = Val(.TextMatrix(i, 5))
    temptable!Val6 = i
    
    temptable!val10 = Val(grid2.TextMatrix(0, 1))
    temptable!val11 = Val(grid2.TextMatrix(0, 3))
    temptable!val12 = Val(grid2.TextMatrix(0, 5))
    temptable!val13 = Val(grid2.TextMatrix(0, 7))
    temptable!VAL14 = Val(grid2.TextMatrix(0, 9))

    temptable!Val16 = Val(grid2.TextMatrix(1, 1))
    temptable!Val17 = Val(grid2.TextMatrix(1, 3))
    temptable!Val20 = Val(grid2.TextMatrix(1, 9))
    temptable!Val21 = Val(grid2.TextMatrix(2, 9))
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\client3.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Sub fixgrd2()
With grid2
.Cell(flexcpBackColor, 0, 0, .Rows - 1, 0) = &H8000000F
.Cell(flexcpBackColor, 0, 2, .Rows - 1, 2) = &H8000000F
.Cell(flexcpBackColor, 0, 4, .Rows - 1, 4) = &H8000000F
.Cell(flexcpBackColor, 0, 6, .Rows - 1, 6) = &H8000000F
.Cell(flexcpBackColor, 0, 8, .Rows - 1, 8) = &H8000000F

.TextMatrix(0, 0) = "—’Ìœ «Ê·"
.TextMatrix(0, 2) = "„»Ì⁄« "
.TextMatrix(0, 4) = "„œ›Ê⁄«  «·Ì"
.TextMatrix(0, 6) = "‘ „—›Ê÷…- ŸÂÌ—"
.TextMatrix(0, 8) = "⁄·ÌÂ"


.TextMatrix(1, 0) = "„ﬁ»Ê÷«  „‰"
.TextMatrix(1, 2) = "œ›⁄«  ‘Ìﬂ« "
.TextMatrix(1, 8) = "·Â"

.TextMatrix(2, 8) = "«·—’Ìœ"
    
.ColWidth(0) = 1400
.ColWidth(1) = 1400
.ColWidth(2) = 1400
.ColWidth(3) = 1400
.ColWidth(4) = 1400
.ColWidth(5) = 1400
.ColWidth(6) = 1600
.ColWidth(7) = 1400
.ColWidth(8) = 1400
.ColWidth(9) = 1400
End With
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub LastOne_GotFocus()
myGotFocus LastOne
End Sub
Private Sub LastOne_LostFocus()
myLostFocus LastOne
End Sub
