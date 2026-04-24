VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PrintedFrm 
   Caption         =   "«” ⁄·«„ ÿ»«⁄…"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   12720
   Begin VB.CommandButton Command2 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      TabIndex        =   43
      Top             =   5580
      Width           =   1860
   End
   Begin VB.CheckBox Check4 
      Caption         =   "·Â„ ’Ê— ›ﬁÿ"
      Height          =   285
      Left            =   90
      TabIndex        =   42
      Top             =   5220
      Width           =   1815
   End
   Begin VB.CheckBox Check3 
      Caption         =   "«⁄÷«¡ ›ﬁÿ"
      Height          =   285
      Left            =   4050
      TabIndex        =   41
      Top             =   6930
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      Caption         =   "«Œ Ì«— «·ﬂ·"
      Height          =   285
      Left            =   90
      TabIndex        =   40
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "«Œ Ì«—«  «·Õ–›"
      Height          =   825
      Left            =   135
      TabIndex        =   38
      Top             =   7245
      Width           =   1860
      Begin VB.CommandButton cmddel 
         Caption         =   "Õ–› «·„Œ «—"
         Height          =   510
         Index           =   0
         Left            =   45
         TabIndex        =   39
         Top             =   225
         Width           =   1725
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   4380
      Left            =   45
      TabIndex        =   32
      Top             =   405
      Width           =   10455
      _cx             =   18441
      _cy             =   7726
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.TextBox xheader 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   45
      Width           =   8925
   End
   Begin VB.Frame Frame2 
      Height          =   2580
      Left            =   5940
      TabIndex        =   23
      Top             =   4815
      Width           =   3915
      Begin VB.TextBox xUsername 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2100
         Width           =   2415
      End
      Begin VB.TextBox xDesca2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1725
         Width           =   2415
      End
      Begin VB.TextBox xDesca1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1350
         Width           =   2415
      End
      Begin VB.ComboBox xType 
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   2565
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   975
         Width           =   1140
      End
      Begin VB.TextBox xCode1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label Label1 
         Caption         =   "«·„” Œœ„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2175
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "«· «»⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ «·⁄÷Ê :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "‰Ê⁄ «·⁄÷ÊÌ… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "≈·Ì —ﬁ„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1050
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "„‰ —ﬁ„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   2775
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   675
         Width           =   690
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   90
      TabIndex        =   17
      Top             =   45
      Width           =   1440
   End
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   390
      Left            =   4500
      Top             =   4350
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "«” ⁄·«„  «—ÌŒ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   4770
      Width           =   3915
      Begin VB.ComboBox xTiming2 
         Height          =   315
         ItemData        =   "Printed1.frx":0000
         Left            =   150
         List            =   "Printed1.frx":000A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   1350
         Width           =   1365
      End
      Begin VB.ComboBox xTiming1 
         Height          =   315
         ItemData        =   "Printed1.frx":001C
         Left            =   150
         List            =   "Printed1.frx":0026
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox xMinute2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2325
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   390
      End
      Begin VB.TextBox xMinute1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2325
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   975
         Width           =   390
      End
      Begin VB.TextBox xYear 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1725
         Width           =   1140
      End
      Begin VB.TextBox xHour1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1575
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   975
         Width           =   465
      End
      Begin VB.TextBox xHour2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1575
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   465
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   1140
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label Label3 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2175
         TabIndex        =   22
         Top             =   1050
         Width           =   165
      End
      Begin VB.Label Label2 
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   21
         Top             =   1425
         Width           =   90
      End
      Begin VB.Label Label1 
         Caption         =   "„Ê”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1725
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "„‰ «·”«⁄… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   975
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "≈·Ì :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "≈·Ì :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   675
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   390
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3690
      _ExtentX        =   6509
      _ExtentY        =   688
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   1980
      TabIndex        =   33
      Top             =   7335
      Width           =   7935
      Begin VB.CommandButton CmdQry 
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3090
         TabIndex        =   36
         Top             =   180
         Width           =   1215
      End
      Begin VB.CheckBox Check1 
         Caption         =   " Ã„Ì⁄ «·«⁄÷«¡"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   225
         Width           =   1440
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4335
         TabIndex        =   34
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label xRecord 
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
         Height          =   405
         Left            =   90
         TabIndex        =   37
         Top             =   180
         Width           =   2940
      End
   End
End
Attribute VB_Name = "PrintedFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
cmddel(0).Enabled = Check1.Value = 0
'CmdDel(1).Enabled = Check1.Value = 0
End Sub

Private Sub Check2_Click()
For i = 1 To Grid1.Rows - 1
    Grid1.TextMatrix(i, 9) = IIf(Check2.Value = 0, "", "-1")
Next
End Sub

Private Sub CmdDel_Click(Index As Integer)
If MsgBox(" Õ–› " & IIf(Index = 0, "«·”Ã·«  «·„Œ «—…", "ﬂ· «·”Ã·« ") & " Â· «‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·« ") = vbYes Then
    If mydelall(Index) Then
        MsgBox " „ «·Õ–› »‰Ã«Õ"
        If Index = 1 Then
            Grid1.Rows = 1
        Else
            fillgrd
        End If
    End If
End If
End Sub

Private Sub Command1_Click()
Load PrintGrd
PrintGrd.doprint Grid1, 1.5, False, xheader.Text, , , , , , , xRecord.Caption
PrintGrd.Show 1
End Sub
Sub fillgrd()
Dim fs As New FileSystemObject
Grid1.Rows = 1

If Check1.Value = 0 Then
    cString = "Select Flag,AUTO,iif(Flag = 1, '⁄÷Ê ⁄«„·','⁄÷Ê ‘—›Ì') as Desca,[Member],desca1,Desca2,Format([Date],'dd-mm-yy') as MYDATE,Format([Time],'hh:mm:AM/PM') as MYTIME,username from QFILE4"
Else
    cString = "Select Flag,First(AUTO),iif(Flag = 1, '⁄÷Ê ⁄«„·','⁄÷Ê ‘—›Ì') as  Desca,[Member],desca1 ,Desca2 ,Format(Max([Date]),'dd-mm-yy') as MYDATE,Format(Max([Time]),'hh:mm:AM/PM') as MYTIME,username from Qfile4"
End If

If xType.ListIndex = 1 Then
    cString = cString & " Where Flag  = 1 "
End If

If xType.ListIndex = 2 Then
    cString = cString & " Where Flag  = 2 "
End If

If xCode1.Text <> "" Then
    cString = cString & turnFound(cString) & " member " & IIf(xCode2.Text = "", "=", " >= ") & Val(xCode1.Text)
End If

If xCode2.Text <> "" Then
    cString = cString & turnFound(cString) & " member <= " & Val(xCode2.Text)
End If

If xDesca1.Text <> "" Then
    cString = cString & turnFound(cString) & " Desca1 like " & MyParn("%" & xDesca1.Text & "%")
End If

If xDesca2.Text <> "" Then
    cString = cString & turnFound(cString) & " Desca2 like " & MyParn("%" & xDesca2.Text & "%")
End If

If xUsername.Text <> "" Then
    cString = cString & turnFound(cString) & " username like " & MyParn("%" & crypt(xUsername.Text) & "%")
End If

If Check3.Value <> 0 Then
    cString = cString & turnFound(cString) & " isNull(Relation) "
End If

If Val(xYear.Text) > 1900 Then cString = cString & addand(cString) & " [season] = " & xYear.Text
If IsDate(xDate1.Text) Then cString = cString & addand(cString) & " [Date] >= " & DateSql(xDate1.Text)
If IsDate(xDate2.Text) Then cString = cString & addand(cString) & " [Date] <= " & DateSql(xDate2.Text)
If Val(xHour1.Text) > 0 Then cString = cString & addand(cString) & " [time] >= " & retTime(xHour1.Text, xMinute1.Text, xTiming1.ListIndex)
If Val(xHour2.Text) > 0 Then cString = cString & addand(cString) & " [time] <= " & retTime(xHour2.Text, xMinute2.Text, xTiming2.ListIndex)

If Check1.Value = 0 Then
    cString = cString & " Order By [season],[Date],[Time],[Member],code"
Else
    cString = cString & " Group by Season,Member,Desca1,Desca2,Code,USERNAME,Flag Order By [season],Max([Date]),Max([Time]),Member,code"
End If
data1.RecordSource = cString
data1.Refresh
If Check4.Value <> 0 Then
    For i = Grid1.Rows - 1 To 1 Step -1
        If Grid1.TextMatrix(i, 0) = 1 Then
            If Not fs.FileExists(RetPhoto(Grid1.TextMatrix(i, 3))) Then Grid1.RemoveItem i
        Else
            If Not fs.FileExists(RETPHOTO2(Grid1.TextMatrix(i, 3))) Then Grid1.RemoveItem i
        End If
    Next
End If
setupgrd

'Set GrdTable = mydb.OpenRecordset(cString)
'If GrdTable.RecordCount = 0 Then Exit Sub
'With Grid1
'    Do
'       .AddItem ""
'       .TextMatrix(.Rows - 1, 0) = TurnValue(GrdTable!Member, Null, "")
'       .TextMatrix(.Rows - 1, 1) = TurnValue(GrdTable!desca1, Null, "")
'       .TextMatrix(.Rows - 1, 2) = TurnValue(GrdTable!desca2, Null, "")
'       .TextMatrix(.Rows - 1, 3) = TurnValue(GrdTable!mydate, Null, "")
'       .TextMatrix(.Rows - 1, 4) = TurnValue(GrdTable!mytime, Null, "")
'       .TextMatrix(.Rows - 1, 5) = Crypt(GrdTable!UserName)
'       GrdTable.MoveNext
'    Loop Until GrdTable.EOF
'End With
End Sub
Private Sub CmdDelAll_Click(nType)
If Grid1.Rows <= 1 Then Exit Sub
If MsgBox("”Ì „ «·¬‰ Õ–› ”Ã·«  " & IIf(nType = 0, "«·”Ã·«  «·„Œ «—…", "ﬂ· «·”Ã·« ") & " Â· «‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2, "Õ–› «·”Ã·« ") = vbYes Then
    Do
        If (nType = 0 And Val(Grid1.TextMatrix(i, Grid1.Cols - 1)) <> 0) Or nType = 1 Then
            Grid1.RemoveItem (i)
        End If
    Loop Until Grid1.Rows = 1
End If
'SetupForm
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdQry_Click()
xRecord.Caption = ""
fillgrd
xRecord.Caption = " ⁄œœ «·«⁄÷«¡ : " & Grid1.Rows - 1
End Sub

Private Sub Command2_Click()
Grid1.Cols = Grid1.Cols + 1
Grid1.TextMatrix(0, Grid1.Cols - 1) = "—ﬁ„ ⁄÷ÊÌ… «·‰ﬁ«»…"
Grid1.ColWidth(Grid1.Cols - 1) = 2000
Grid1.ColHidden(2) = True
Grid1.ColHidden(5) = True
Grid1.ColHidden(6) = True
Grid1.ColHidden(7) = True
Grid1.ColHidden(8) = True
Grid1.ColHidden(9) = True


rsMember.Index = "ndxcode"
For i = 1 To Grid1.Rows - 1
    If Grid1.TextMatrix(i, 0) = 1 Then
        rsMember.Seek "=", Grid1.TextMatrix(i, 3)
        If Not rsMember.NoMatch Then
            Grid1.TextMatrix(i, Grid1.Cols - 1) = rsMember!Union & ""
        End If
    End If
Next
PrintGrd.bText = True
Load PrintGrd
PrintGrd.doprint Grid1, 1.5, False, xheader.Text, , , , , , , xRecord.Caption
PrintGrd.Show 1

Grid1.Cols = Grid1.Cols - 1
Grid1.ColHidden(2) = False
Grid1.ColHidden(5) = False
Grid1.ColHidden(6) = False
Grid1.ColHidden(7) = False
Grid1.ColHidden(8) = False
Grid1.ColHidden(9) = False

End Sub

Private Sub Form_Load()
Grid1.ExplorerBar = flexExSort
setupgrd
xTiming1.ListIndex = 0
xTiming2.ListIndex = 0
xType.AddItem "«·ﬂ·"
xType.AddItem "⁄÷Ê ⁄«„·"
xType.AddItem "⁄÷Ê ‘—›Ì"
xType.ListIndex = 0
Check2.Visible = nUserLevel = 10
Frame4.Visible = nUserLevel = 10
'With Ado1
'
'    .Mode = adModeReadWrite
'    Ado1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=D:\Horse\MDB\Data.mdb"
'    Ado1.CommandType = adCmdText
'    Ado1.RecordSource = "Select Member as [—ﬁ„ «·⁄÷Ê],desca as [≈”„ «·⁄÷Ê],Format(Date,'dd-mm-yy') as [«· «—ÌŒ],Format(Time,'hh:mm:AM/PM') as [«·Êﬁ ] from file4_10"
'    Ado1.Refresh
'End With

'SetupForm
'data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\mdb\data.mdb"
'data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\mdb\data.mdb"
Set Grid1.DataSource = data1
data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & MdbPath
data1.CommandType = adCmdText
cString = "Select [Member],desca1,Desca2 from qfile4"

data1.RecordSource = cString
Grid1.Rows = 1
setupgrd
End Sub
Private Sub SetupForm()
If Grid1.Rows = 1 Then
    Frame1.Top = 150
    Me.Height = Frame1.Top + Frame1.Height + CmdExit.Height + 1000
    Grid1.Visible = False
    CmdExit.Top = Frame1.Top + Frame1.Height + 50
Else
    Grid1.Top = 150
    Frame1.Top = 150 + Grid1.Height + 150
    Me.Height = Frame1.Top + Frame1.Height + 450 + CmdExit.Height
    CmdExit.Top = Frame1.Top + Frame1.Height + 50
    Grid1.Visible = True
End If
End Sub
Private Sub Grid1_EnterCell()
Grid1.Editable = IIf(Grid1.Col = 9, flexEDKbdMouse, flexEDNone)
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
'    If Grid1.Row = Grid1.Rows - 1 Then Exit Sub
'   If MsgBox("Õ–› ”Ã· ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        Grid1.RemoveItem Grid1.Row
'    End If
End If
End Sub

Private Sub xTime1_Change()

End Sub
Private Sub xMinute_Validate(Cancel As Boolean)
If Val(xMinute1) > 60 Then
    MsgBox " ÊﬁÌ  €Ì— ’«·Õ"
    Cancel = True
End If
End Sub

Private Sub xHour1_Change()
If Val(xHour1) > 23 Then
    MsgBox " ÊﬁÌ  €Ì— ’«·Õ"
    Cancel = True
End If
End Sub
Private Sub xHour2_Change()
If Val(xHour2) > 23 Then
    MsgBox " ÊﬁÌ  €Ì— ’«·Õ"
    Cancel = True
End If
End Sub
Private Sub xMinute1_Change()
If Val(xMinute2) > 59 Then
    MsgBox " ÊﬁÌ  €Ì— ’«·Õ"
    Cancel = True
End If
End Sub

Private Sub xMinute2_Change()
If Val(xMinute2) > 59 Then
    MsgBox " ÊﬁÌ  €Ì— ’«·Õ"
    Cancel = True
End If
End Sub
Private Function retTime(bHour, bMinute, nTiming)
bHour = String(2 - Len(Trim(bHour)), "0") & Trim(bHour)
bMinute = ":" & String(2 - Len(Trim(bMinute)), "0") & Trim(bMinute)
retTime = "#" & bHour & bMinute & IIf(nTiming = 0, " AM ", " PM ") & "#"
End Function
Private Sub setupgrd()
With Grid1
Grid1.Cols = 10
.TextMatrix(0, 2) = "‰Ê⁄ «·⁄÷ÊÌ…"
.TextMatrix(0, 3) = "—ﬁ„ «·⁄÷Ê"
.TextMatrix(0, 4) = "≈”„ «·⁄÷Ê"
.TextMatrix(0, 5) = "≈”„ «· «»⁄"
.TextMatrix(0, 6) = "«· «—ÌŒ"
.TextMatrix(0, 7) = "«·Êﬁ "
.TextMatrix(0, 8) = "«·„” Œœ„"
.TextMatrix(0, 9) = "Õ–›"
.ColWidth(2) = .Width * 10 / 100 - 100
.ColWidth(3) = .Width * 10 / 100 - 100
.ColWidth(4) = .Width * 20 / 100 - 100
.ColWidth(5) = .Width * 20 / 100 - 100
.ColWidth(6) = .Width * 12 / 100 - 100
.ColWidth(7) = .Width * 12 / 100 - 100
.ColWidth(8) = .Width * 12 / 100 - 100
.ColWidth(9) = .Width * 7 / 100 - 100
.ColHidden(0) = True
.ColHidden(1) = True
.ColDataType(9) = flexDTBoolean
.ColHidden(9) = nUserLevel < 10
For i = 0 To Grid1.Cols - 1
    Grid1.ColAlignment(i) = flexAlignRightCenter
Next
For i = 1 To Grid1.Rows - 1
    .TextMatrix(i, 8) = crypt(Grid1.TextMatrix(i, 8))
Next
End With
End Sub
Private Function mydelall(Index) As Boolean
On Error GoTo myError:
Myws.BeginTrans
For i = 1 To Grid1.Rows - 1
    If (Val(Grid1.TextMatrix(i, 9)) <> 0 And Index = 0) Or Index = 1 Then
        If Grid1.TextMatrix(i, 0) = "1" Then
            mydb.Execute "Delete * from file4_10 where auto = " & Grid1.TextMatrix(i, 1)
        Else
            mydb.Execute "Delete * from file4_20 where auto = " & Grid1.TextMatrix(i, 1)
        End If
    End If
Next
Myws.CommitTrans
mydelall = True
Exit Function
myError:
Myws.Rollback
MsgBox Err.Description
Err.Clear
End Function

