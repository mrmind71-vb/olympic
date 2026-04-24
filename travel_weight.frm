VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form travel_weightfrm 
   Caption         =   "«·«Ê“«‰"
   ClientHeight    =   5475
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   11700
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5130
      MaskColor       =   &H00FFFFFF&
      Picture         =   "travel_weight.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Õ›Ÿ"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   1635
   End
   Begin VB.CommandButton CmdUndo 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   3450
      MaskColor       =   &H00FFFFFF&
      Picture         =   "travel_weight.frx":2363
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   " —«Ã⁄"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   1680
   End
   Begin VB.CommandButton CmdDel 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   1755
      MaskColor       =   &H00FFFFFF&
      Picture         =   "travel_weight.frx":48DC
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Õ–›"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   1680
   End
   Begin VB.Frame Frame2 
      Height          =   3210
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2205
      Width           =   4740
      Begin VB.TextBox xExtend 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1260
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2430
         Width           =   2040
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1260
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2070
         Width           =   2040
      End
      Begin VB.Label Label23 
         Caption         =   "«·“Ì«œ…"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2475
         Width           =   1050
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2790
         Width           =   3165
      End
      Begin VB.Label Label21 
         Caption         =   "«·‰Ê·Ê‰"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2835
         Width           =   1050
      End
      Begin VB.Label Label20 
         Caption         =   "«·Œ’„"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2115
         Width           =   1050
      End
      Begin VB.Label xWeight_Value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1350
         Width           =   3165
      End
      Begin VB.Label Label18 
         Caption         =   "ﬁÌ„… «·Ê“‰"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1395
         Width           =   1050
      End
      Begin VB.Label xDiffer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1710
         Width           =   3165
      End
      Begin VB.Label Label16 
         Caption         =   "«·›—ﬁ"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label xWeight_Total 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   3165
      End
      Begin VB.Label Label14 
         Caption         =   "Ê“‰ «·‰ﬁ·…"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   1050
      End
      Begin VB.Label xStand_Value 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   3165
      End
      Begin VB.Label Label12 
         Caption         =   "ﬁÌ„… «·Õœ «·«œ‰Ì"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label xStand_weight 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   3165
      End
      Begin VB.Label Label10 
         Caption         =   "Õœ Ê“‰ «œ‰Ì"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "»Ì«‰«  «·”›—Ì…"
      Height          =   2220
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   4785
      Begin VB.Label xTrailer 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1755
         Width           =   3165
      End
      Begin VB.Label xDoc_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1260
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label8 
         Caption         =   "‰Ê⁄ «·„ﬁÿÊ—…"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   1185
      End
      Begin VB.Label xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   3165
      End
      Begin VB.Label Label6 
         Caption         =   "«· «—ÌŒ"
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   555
      End
      Begin VB.Label xPlace2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1395
         Width           =   3165
      End
      Begin VB.Label Label4 
         Caption         =   "≈·Ì"
         Height          =   330
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label xPlace1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1035
         Width           =   3165
      End
      Begin VB.Label Label2 
         Caption         =   "„‰"
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   555
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   315
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "«·⁄„Ì·"
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      Picture         =   "travel_weight.frx":7176
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   1680
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   810
      Top             =   855
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
   Begin MSAdodcLib.Adodc data11 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4650
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6675
      _cx             =   11774
      _cy             =   8202
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   12
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
      FixedRows       =   1
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
Attribute VB_Name = "travel_weightfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Public sDoc_no As String
Const LoadMode = 1, DefineMode = 2
Private Sub myload()
Dim cString As String
If Not loctable.EOF Then
    xDesca.Caption = loctable!desca & ""
    xDate.Caption = Format(loctable!Date, "yyyy/m/d")
    xPlace1.Caption = loctable!Place_desca1 & ""
    xPlace2.Caption = loctable!place_desca2 & ""
    xTrailer.Caption = loctable!Trailer_desca & ""
    xStand_weight.Caption = Myvalue(loctable!STAND_weight)
    xStand_Value.Caption = Myvalue(loctable!STAND_Value)
    xExtend.Text = Myvalue(loctable!Extend)
    xDiscount.Text = Myvalue(loctable!Discount)
    If (Not IsNull(loctable!trailer)) And (Not IsNull(loctable!place1)) And (Not IsNull(loctable!place2)) Then
        Dim aRet As Variant
        cString = "select fair.weight,fair_sub.[value] from fair inner join fair_sub on fair.code = fair_sub.code "
        cString = cString & " where fair.Trailer = " & loctable!trailer
        cString = cString & " and (fair_sub.place1 = " & loctable!place1 & " and fair_sub.place2 = " & loctable!place2 & ")"
'        cString = cString & " and (fair_sub.place1 = " & loctable!place1 & " Or fair_sub.place2 = " & loctable!place1 & ")"
'        cString = cString & " and (fair_sub.place1 = " & loctable!place2 & " Or fair_sub.place2 = " & loctable!place2 & ")"
        cString = cString & " and fair.client = " & MyParn(loctable!code)
        aRet = GetFields(cString, con)
        If Not IsEmpty(aRet) Then
            If xStand_weight.Caption = "" Then xStand_weight.Caption = Myvalue(retFlag(aRet, "Weight"))
            If xStand_Value.Caption = "" Then xStand_Value.Caption = Myvalue(retFlag(aRet, "value"))
        End If
    End If
End If
myloadgrd
'Set data1.Recordset = myRecordSet(cString, con)
'myAddItem
'Fixgrd
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
con.BeginTrans
con.Execute "delete from travel_w where doc_no = " & MyParn(sDoc_no)
con.CommitTrans
myloadgrd
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If myreplace Then
    Inform " „ «·Õ›Ÿ »‰Ã«Õ"
    Unload Me
End If
End Sub
Private Sub CmdUndo_Click()
If MsgBox("«⁄«œ… ÷»ÿ Õ”«» «·‰Ê·Ê‰", vbOKCancel + vbDefaultButton2) Then
    If (Not IsNull(loctable!trailer)) And (Not IsNull(loctable!place1)) And (Not IsNull(loctable!place2)) Then
        Dim aRet As Variant
        cString = "select fair.weight,fair_sub.[value] from fair inner join fair_sub on fair.code = fair_sub.code "
        cString = cString & " where fair.Trailer = " & loctable!trailer
        cString = cString & " and (fair_sub.place1 = " & loctable!place1 & " Or fair_sub.place2 = " & loctable!place1 & ")"
        cString = cString & " and (fair_sub.place1 = " & loctable!place2 & " Or fair_sub.place2 = " & loctable!place2 & ")"
        cString = cString & " and fair.client = " & MyParn(loctable!code)
        aRet = GetFields(cString, con)
        xStand_weight.Caption = Myvalue(retFlag(aRet, "Weight"))
        xStand_Value.Caption = Myvalue(retFlag(aRet, "value"))
        If myreplace Then Inform " „ «⁄«œ… Õ”«» «·‰Ê·Ê‰ »‰Ã«Õ"
    End If
End If
End Sub
Private Sub Form_Activate()
If loctable.EOF And loctable.BOF Then
    MsgBox "«·»Ì«‰«  ·Ì”  ﬂ«›Ì… ·Õ”«» «·‰Ê·Ê‰"
    Unload Me
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
loctable.Close
Set loctable = Nothing
Set containerfrm = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Calctotals
On Error GoTo myerror
myreplace Row
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Or (Val(grid1.TextMatrix(Row, 3)) = 0 And grid1.TextMatrix(Row, grid1.Cols - 1) <> "") Then
    myloadgrd
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Grid1_EnterCell()
If (grid1.Col = 3) Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data11

cString = "SELECT TRAVEL_H.DOC_NO, TRAVEL_H.DATE, TRAVEL_H.STAND_WEIGHT, TRAVEL_H.STAND_VALUE, TRAVEL_H.DISCOUNT, TRAVEL_H.EXTEND,PLACE_CODES.DESCA AS PLACE_DESCA1,PLACE_CODES_1.DESCA AS PLACE_DESCA2,TRAILER_CODES.DESCA AS TRAILER_DESCA" & _
           ",TRAVEL_H.PLACE1,TRAVEL_H.PLACE2,TRAVEL_H.TRAILER" & _
           ", TRAVEL_H.CODE, FILE3_10.DESCA AS DESCA" & _
           " FROM  TRAVEL_H INNER JOIN PLACE_CODES ON TRAVEL_H.PLACE1 = PLACE_CODES.CODE INNER JOIN" & _
           " PLACE_CODES AS PLACE_CODES_1 ON TRAVEL_H.PLACE2 = PLACE_CODES_1.CODE INNER JOIN " & _
           "  TRAILER_CODES ON TRAVEL_H.TRAILER = TRAILER_CODES.CODE" & _
           " INNER JOIN FILE3_10 ON TRAVEL_H.CODE = FILE3_10.CODE" & _
           " INNER JOIN FAIR ON (FAIR.CLIENT = TRAVEL_H.CODE AND FAIR.TRAILER = TRAVEL_H.TRAILER)" & _
           " INNER JOIN FAIR_SUB ON (FAIR.CODE = FAIR_SUB.CODE AND FAIR_SUB.PLACE1 = TRAVEL_H.PLACE1 AND FAIR_SUB.PLACE2 = TRAVEL_H.PLACE2)"
cString = cString & turn(cString) & "TRAVEL_H.DOC_NO = " & MyParn(sDoc_no)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
'With grid1
myload
CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Row > 0 Then
    If MsgBox("Õ–› !! Â· «‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2) = vbYes Then
        If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
            con.BeginTrans
            On Error GoTo myerror
            con.Execute "Delete from Container where code = " & grid1.TextMatrix(grid1.Row, 0)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myload
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "«·ﬂÊœ „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
ElseIf Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "«·«”„ „ﬂ Ê»"
        Cancel = True
    End If
ElseIf Col = 2 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "«·Ê“‰ „ÿ·Ê»"
        Cancel = True
    End If
End If
End Sub
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 600
.ColWidth(1) = 2000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColHidden(0) = True
.ColHidden(.Cols - 1) = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xDesca_Change()
myload
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .Rows - 1 And .Row <> 0 And .TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    .RemoveItem .Row
End If
End With
End Sub
Private Function validRow(Row) As Boolean
'If Trim(grid1.TextMatrix(Row, 1)) = "" Then Exit Function
'If Trim(grid1.TextMatrix(Row, 2)) = "" Then Exit Function
validRow = True
End Function
Private Sub MyAddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 3 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 3, 3)
    grid1.ShowCell grid1.Row, 0
ElseIf Row = grid1.Rows - 1 And Col >= grid1.Cols - 3 Then
    cmdSave.SetFocus
End If
End Sub

Private Sub Label7_Click()

End Sub
Private Sub myloadgrd()
With grid1
Dim cString As String
cString = "SELECT CONTAINER.CODE AS [«·ﬂÊœ],CONTAINER.DESCA AS [«·⁄»Ê…],CONTAINER.WEIGHT AS [«·Ê“‰],TRAVEL_W.QUANT AS [«·⁄œœ],CONTAINER.WEIGHT * TRAVEL_W.QUANT AS [≈Ã„«·Ì «·Ê“‰],TRAVEL_W.ID " & _
          " FROM CONTAINER LEFT JOIN TRAVEL_W ON CONTAINER.CODE = TRAVEL_W.CONTAINER AND TRAVEL_W.DOC_NO = " & MyParn(sDoc_no)
Set data11.Recordset = myRecordSet(cString, con)
End With
Fixgrd
Calctotals
End Sub
Private Function myreplace(Optional Row As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(aInsert, "[STAND_WEIGHT]", Val(xStand_weight.Caption))
aInsert = AddFlag(aInsert, "[STAND_VALUE]", Val(xStand_Value.Caption))
aInsert = AddFlag(aInsert, "[WEIGHT_TOTAL]", Val(xWeight_Total.Caption))
aInsert = AddFlag(aInsert, "[WEIGHT_VALUE]", Val(xWeight_Value.Caption))
aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(xDiscount.Text))
aInsert = AddFlag(aInsert, "[EXTEND]", Val(xExtend.Text))
aInsert = AddFlag(aInsert, "[TOTAL]", Val(xTotal.Caption))
aInsert = AddFlag(aInsert, "[WEIGHT]", "1")
aInsert = AddFlag(aInsert, "[CLASS]", Val(xTotal.Caption))

'aInsert = AddFlag(aInsert, "[WEIGHT]", Val(xWeight_Total.Caption))
'If Val(xWeight_Total.Caption) <> 0 Then
'    aInsert = AddFlag(aInsert, "[CLASS]", Round(Val(xTotal.Caption) / Val(xWeight_Total.Caption), 6))
'Else
'    aInsert = AddFlag(aInsert, "[CLASS]", "0")
'End If
On Error GoTo myerror
con.BeginTrans
con.Execute addUpdate(aInsert, "TRAVEL_H", "DOC_NO = " & MyParn(sDoc_no))
myreplaceGrd Row
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub Calctotals()
Dim nWeight_Total As Double
With grid1
For i = 1 To .Rows - 1
    .TextMatrix(i, 4) = Myvalue(Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 3)))
    nWeight_Total = nWeight_Total + Val(.TextMatrix(i, 4))
Next
xWeight_Total.Caption = nWeight_Total
If Val(xStand_weight.Caption) <> 0 Then
    xWeight_Value.Caption = Round(IIf(Val(xWeight_Total.Caption) / Val(xStand_weight.Caption) > 1, Val(xWeight_Total.Caption) / Val(xStand_weight.Caption), 1) * Val(xStand_Value.Caption), 2)
Else
    xWeight_Value.Caption = ""
End If
xDiffer.Caption = Val(xWeight_Value.Caption) - Val(Me.xStand_Value.Caption)
If Val(xDiffer.Caption) <> 0 Then xDiffer.Caption = IIf(Val(xDiffer.Caption) > 0, "+ ", "- ") & Val(xDiffer.Caption)
xTotal.Caption = Val(xWeight_Value.Caption) + Val(xExtend.Text) - Val(xDiscount.Text)
End With
End Sub
Private Function myreplaceGrd(Row) As Boolean
With grid1
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, .Rows - 1, Row)
        If Val(.TextMatrix(i, 3)) = 0 And .TextMatrix(i, .Cols - 1) <> "" Then
            con.Execute "DELETE FROM TRAVEL_W WHERE ID = " & .TextMatrix(Row, .Cols - 1)
        ElseIf Val(.TextMatrix(i, 3)) <> 0 Then
            Dim aInsert As Variant
            aInsert = AddFlag(Empty, "DOC_NO", addstring(sDoc_no))
            aInsert = AddFlag(aInsert, "CONTAINER", Val(.TextMatrix(i, 0)))
            aInsert = AddFlag(aInsert, "[QUANT]", Val(.TextMatrix(i, 3)))
            aInsert = AddFlag(aInsert, "[TOTAL]", Val(.TextMatrix(i, 4)))
            If .TextMatrix(i, .Cols - 1) = "" Then
                con.Execute addInsert(aInsert, "TRAVEL_W")
            Else
                con.Execute addUpdate(aInsert, "TRAVEL_W", "ID = " & .TextMatrix(i, .Cols - 1))
            End If
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub xDiffer_Change()
xDiffer.ForeColor = IIf(Val(xDiffer.Caption) > 0, vbRed, &H80000008)
grid1.BackColorFixed = IIf(Val(xDiffer.Caption) > 0, &HC0FFC0, &H80000010)
End Sub
Private Sub xDiscount_LostFocus()
Calctotals
myreplace
End Sub
Private Sub xExtend_LostFocus()
Calctotals
myreplace
End Sub
