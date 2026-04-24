VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form PurPayfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "”œ«œ «·›« Ê—…"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "≈Ã„«·Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1935
      Width           =   5280
      Begin VB.Label xRestCur 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1530
         Width           =   3120
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " ’›Ì… „⁄ «· ⁄œÌ· :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1620
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„ »ﬁÌ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1215
         Width           =   600
      End
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   3120
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈Ã„«·Ì «·›« Ê—… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   3120
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„œ›Ê⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   765
         Width           =   705
      End
      Begin VB.Label xPaid 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   3120
      End
   End
   Begin VB.Frame FrameBefore 
      Caption         =   "”œ«œ ”«»ﬁ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3915
      Width           =   5280
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„»·€ «·„”œœ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1440
         Width           =   1185
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " «—ÌŒ «·„” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   675
         Width           =   1155
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label xdoc_no 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   225
         Width           =   3165
      End
      Begin VB.Label xDoc_Box 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   3165
      End
      Begin VB.Label xDoc_Value 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1350
         Width           =   3165
      End
      Begin VB.Label xDoc_Date 
         Alignment       =   1  'Right Justify
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
         Height          =   330
         Left            =   585
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   3165
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "”œ«œ ÃœÌœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   5295
      Begin VB.TextBox xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox xvalue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   45
         TabIndex        =   10
         Top             =   945
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblPaid 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "«·›« Ê—… „”œœ… »«·ﬂ«„·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   225
         Width           =   2130
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„»·€ «·„ »ﬁÌ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1350
      Width           =   5280
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–›"
         Height          =   420
         Left            =   1260
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   " ⁄œÌ·"
         Height          =   420
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "√÷«›…"
         Height          =   420
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   135
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin VB.Frame Frame8 
      Height          =   555
      Left            =   3105
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5715
      Width           =   1905
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin VB.Label lblRecord 
      Height          =   375
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   5850
      Width           =   2850
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1215
      Width           =   45
   End
End
Attribute VB_Name = "PurPayfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As New ADODB.Recordset, nOrgHeight
Private Sub CmdApply_Click()
'con.Execute "Delete * from file8_10h where doc_no = " & MyParn(cDoc_No)
'con.Execute "Delete * from file8_10 where doc_no = " & MyParn(cDoc_No)
If MyReplace Then
    MsgBox " „ «·Õ›Ÿ"
    'Doprint
End If
Unload Me

End Sub

Private Sub CmdAdd_Click()
If Val(xValue.Text) <> Val(xRest.Caption) Then
    If MsgBox("ﬁÌ„… «·›« Ê—… Ê«·”œ«œ €Ì— „ ”«ÊÌ…  ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
End If
'If xdoc_no.Caption <> "" And (Val(xvalue.Text) <> Val(xRest.Caption)) Then
'    If MsgBox("Â‰«ﬂ „” ‰œ ”œ«œ »«·›⁄· ··›« Ê—… .. «÷«›… „” ‰œ «Œ— ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
'End If
If MyReplace Then
    CardTable.Requery
    CardTable.MoveLast
    MyLoad
    xValue.Text = xRest.Caption
End If
End Sub

Private Sub CmdDel_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) = vbOK Then
    If myDel Then
        CardTable.Requery
        CardTable.Find "doc_no < " & MyParn(xDoc_No.Caption), , adSearchBackward, adBookmarkLast
        If CardTable.BOF And Not (CardTable.EOF) Then CardTable.MoveFirst
        MyLoad
        xValue.Text = xRest.Caption
    End If
End If
End Sub

Private Sub cmdEdit_Click()
If Val(xValue.Text) - Val(xDoc_Value.Caption) <> Val(xRest.Caption) Then
    If MsgBox("ﬁÌ„… «·›« Ê—… Ê«·”œ«œ €Ì— „ ”«ÊÌ…  ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
End If
If MyEdit Then
    MyLoad
    xValue.Text = xRest.Caption
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdUndo_Click()
If MyReplace Then
    MsgBox " „ «·Õ›Ÿ"
    'doprint
End If
Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
    
MYVALID = True
End Function
Private Sub doprint3()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(3)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA AS ITEMDESCA,FILE1_10.UNIT, " & _
          "  FILE7_20.DOC_NO,FILE7_21.CODE as SupCode,FILE7_21.DATE,FILE7_21.VESSEL, " & _
          " Sum(val(FILE7_20.Quant & '')) AS SumofQuant, " & _
          " FILE1_10.[GROUP] AS GroupCode, FILE1_50.DESCA AS GroupDesca,  " & _
          " FILE1_50.[GROUP] AS MainGroupCode, FILE1_51.DESCA as  MainGroupDesca" & _
          " FROM (((FILE7_20 INNER JOIN FILE7_21 ON FILE7_20.DOC_NO = FILE7_21.DOC_NO )INNER JOIN FILE1_10  ON FILE7_20.ITEM = FILE1_10.ITEM) LEFT " & _
          " JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_51 ON " & _
          " FILE1_50.[GROUP] = FILE1_51.CODE"

If XCODE.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.CODE = " & MyParn(XCODE.Text)
    aHeader(0) = "[" & "Supplier :" & xCodeDesca.Caption & "]"
End If

If xDoc_No.Caption <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.doc_no = " & MyParn(xDoc_No.Caption)
    aHeader(1) = "[" & "Inovice No." & xDoc_No.Caption & "]"
End If

If xdate1.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.date >= " & DATESQ(xdate1.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
End If

If XDATE2.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.date <= " & DATESQ(XDATE2.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
End If

cString = cString & " Group by  FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA,FILE1_10.UNIT," & _
          " FILE7_20.DOC_NO,FILE7_21.CODE,FILE7_21.DATE,FILE7_21.VESSEL, " & _
          " FILE1_10.[GROUP], FILE1_50.DESCA,FILE1_50.[GROUP], FILE1_51.DESCA  "

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!STR6 = !MainGroupDesca
        temptable!str5 = !MAINGROUPCODE
        temptable!str1 = !Item
        temptable!str2 = itemWidth(sourcetable!Item)
        temptable!str3 = !GroupCode
        temptable!str4 = !GroupDesca
        temptable!str8 = GetDesca("Select Desca from file1_13 where code = " & MyParn(!UNIT))
        temptable!str9 = !doc_no
        temptable!str10 = GetDesca("Select Desca from file4_10 where code = " & MyParn(!supCode))
        temptable!Str11 = !Vessel
        temptable!Date1 = !Date
        temptable!val1 = !sumOfQuant
        temptable!Val20 = !width1
        temptable!STR20 = !width1
        temptable!str17 = TurnValue(retHeader(aHeader, 0, 4))
        temptable.Update
        .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.REPORT1.ReportFileName = App.Path & "\Reports\Item3.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub

Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then DocLookup
End Sub
Private Function itemWidth(pItem) As String
itemWidth = retitem(pItem, "width1") & ""
If Not IsNull(retitem(pItem, "width2")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "width2")
If Not IsNull(retitem(pItem, "length")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "length")
End Function
Private Sub doprint4()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(3)

contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT Sum(Val(FILE7_20.[QUANT] & '')) AS SumOfQuant, Sum(Val(FILE7_20.[QUANT] & '')* VAL(FILE1_10.PRICE & '')) AS SumOfValue, FILE1_10.[GROUP] AS GroupCode, FILE1_50.DESCA AS GroupDesca, FILE1_50.[GROUP] AS MainGroupCode, FILE1_51.DESCA AS MainGroupDesca, FILE1_13.DESCA AS UNITDESCA,FILE7_20.DOC_NO,FILE7_21.DATE,FILE7_21.CODE,FILE7_21.DESCA AS SUPDESCA " & _
          " FROM (((((FILE7_20 INNER JOIN FILE7_21 ON FILE7_20.DOC_NO = FILE7_21.DOC_NO) INNER JOIN FILE4_10 ON FILE7_21.CODE = FILE4_10.CODE)INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM) LEFT JOIN FILE1_13 ON FILE1_10.UNIT = FILE1_13.CODE) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_51 ON FILE1_50.[GROUP] = FILE1_51.CODE"

If XCODE.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.CODE = " & MyParn(XCODE.Text)
    aHeader(0) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If

If xDoc_No.Caption <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.doc_no = " & MyParn(xDoc_No.Caption)
    aHeader(1) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If

If xdate1.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.date >= " & DATESQ(xdate1.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
End If

If XDATE2.Text <> "" Then
    cString = cString & turnFound2(cString) & "File7_21.date <= " & DATESQ(XDATE2.Text)
    aHeader(2) = "[" & BetweenString(xdate1.Text, XDATE2.Text) & "]"
End If

cString = "Group by FILE1_10.[GROUP], FILE1_50.DESCA, FILE1_50.[GROUP] , FILE1_51.DESCA, FILE1_13.DESCA AS UNITDESCA,FILE7_20.DOC_NO,FILE7_21.DATE,FILE7_21.CODE,FILE7_21.DESCA AS SUPDESCA "

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        cCondition = IIf(xShowMinus.Value = 0, sourcetable!balance > 0, sourcetable!balance < 0)
        If cCondition Then
            temptable.AddNew
            temptable!STR6 = !MainGroupDesca
            temptable!str5 = !MAINGROUPCODE
            temptable!str1 = !GroupCode
            temptable!str2 = !GroupDesca
            temptable!str8 = !unitDesca
            temptable!val1 = Val(!balance & "")
            temptable!val2 = Val(!BALANCEVALUE & "")
            temptable!str7 = "√—’œ… «·√’‰«›"
            temptable.Update
        End If
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    If xShowCost.Value = 0 Then
       main.REPORT1.ReportFileName = App.Path & "\Reports\Item1.rpt"
    Else
        main.REPORT1.ReportFileName = App.Path & "\Reports\item1_2.rpt"
    End If
    
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Load()
nOrgHeight = Me.Height
'xdoc_no.caption = RetZero(Val(GetDesca("Select Max(doc_no) from file8_10h")) + 1)
CardTable.Open "select file8_20H.doc_no,file8_20h.date,file8_20.value,file0_50.desca from (file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no) Left join file0_50 on file8_20.box = file0_50.code where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text) & " Order by file8_20h.Doc_no Desc", con, adOpenKeyset, adLockReadOnly, adCmdText
data1.ConnectionString = con.ConnectionString
data1.RecordSource = "Select * From file0_50"

Set XBOX.RowSource = data1
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

If Not (data1.Recordset.EOF And data1.Recordset.BOF) Then
    data1.Recordset.MoveFirst
    XBOX.BoundText = data1.Recordset!CODE
End If
xValue.Text = Purchasefrm.xTotal.Caption
xDate.Text = Purchasefrm.xDate.Text
MyLoad
xValue.Text = xRest.Caption
End Sub
Private Function MyReplace() As Boolean
cDoc_no = RetZero(Val(GetDesca("Select Max(doc_no) from file8_20h")) + 1)
On Error Resume Next
For i = 1 To 10
    con.BeginTrans
    con.Execute "insert into file8_20h(doc_no,[date],Doc_No_Pur)" & _
              " Values(" & _
              addstring(cDoc_no) & "," & _
              DATESQ(xDate.Text) & "," & _
              addstring(Purchasefrm.xDoc_No.Text) & _
              ")"
    If Err.Number = 0 Then
        con.Execute "Insert Into file8_20(Doc_No,[Date],Code,Desca,[Value],Box,Row,username) " & _
                  " Values(" & _
                  addstring(cDoc_no) & "," & _
                  DATESQ(xDate.Text) & "," & _
                  addstring(Purchasefrm.XCODE.Text) & "," & _
                  addstring("”œ«œ ›« Ê—… „‘ —Ì«  —ﬁ„ " & Format(Purchasefrm.xDoc_No.Text) & " » «—ÌŒ : " & Purchasefrm.xDate.Text) & "," & _
                  Val(xValue.Text) & "," & _
                  addstring(XBOX.BoundText) & "," & _
                  i & "," & _
                  addstring(sUserName) & _
                  ")"
        If Err.Number <> 0 Then GoTo myerror
    End If
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 Then
        cDoc_no = RetZero(Val(cDoc_no) + 1)
        Err.Clear
        con.RollbackTrans
    End If
    If Err.Number <> 0 Then GoTo myerror
Next
con.CommitTrans
MyReplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Function MyEdit() As Boolean
'On Error GoTo MYERROR
con.BeginTrans
con.Execute "update file8_20 SET FILE8_20.VALUE = " & Val(xValue.Text) & "," & _
            " file8_20.box = " & addstring(XBOX.BoundText) & _
            " WHERE DOC_NO = " & MyParn(xDoc_No.Caption)
con.Execute "UPDATE file8_20h SET FILE8_20H.DATE = " & DATESQ(xDate.Text) & _
            " where file8_20h.doc_no = " & MyParn(xDoc_No.Caption)
con.CommitTrans
MyEdit = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub MyLoad()
With CardTable
If Not (.EOF And .BOF) Then
    nBookMark = CardTable.Bookmark
    CardTable.MoveLast
    nRecordCount = CardTable.RecordCount
    CardTable.Bookmark = nBookMark
    'lblRecord.Caption = " „” ‰œ  " & (CARDTABLE.AbsolutePosition + 1) & " „‰ " & nRecordcount
    lblRecord.Caption = "⁄œœ «·”Ã·«  " & nRecordCount
    cmdfirst.Enabled = True
    cmdLast.Enabled = True
    cmdNext.Enabled = True
    cmdPrevious.Enabled = True
    xDoc_No.Caption = !doc_no
    xDoc_Date.Caption = Format(!Date, "dd-mm-yyyy")
    xDoc_Box.Caption = !DESCA & ""
    xDoc_Value.Caption = Format(Val(!Value & ""), "Fixed")
    cmdEdit.Enabled = True
    CmdDel.Enabled = True
    FrameBefore.Visible = True
    Frame8.Visible = True
    Frame3.Visible = True
    Me.Height = nOrgHeight
Else
    cmdfirst.Enabled = False
    cmdLast.Enabled = False
    cmdNext.Enabled = False
    cmdPrevious.Enabled = False
    xDoc_No.Caption = ""
    xDoc_Date.Caption = ""
    xDoc_Box.Caption = ""
    xDoc_Value.Caption = ""
    FrameBefore.Visible = False
    Me.Height = 2565
    cmdEdit.Enabled = False
    CmdDel.Enabled = False
    Frame8.Visible = False
    Frame3.Visible = False
    lblRecord.Caption = ""
End If
End With
xTotal.Caption = Format(Val(Purchasefrm.xTotal.Caption))
xPaid.Caption = Format(Val(GetDesca("Select Sum(value) From file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text))), "Fixed")
xRest.Caption = Format(Val(Purchasefrm.xTotal.Caption) - Val(GetDesca("Select Sum(value) From file8_20 inner join file8_20h on file8_20.doc_no = file8_20h.doc_no where file8_20h.doc_no_pur = " & MyParn(Purchasefrm.xDoc_No.Text))) & "", "Fixed")
xRestCur.Caption = Format(Val(xRest.Caption) + Val(xDoc_Value.Caption), "Fixed")
xRest.ForeColor = IIf(Val(xRest.Caption) = 0, vbBlack, vbRed)
lblPaid.Visible = Val(xRest.Caption) <= 0
End Sub
Private Function myDel() As Boolean
con.BeginTrans
con.Execute "Delete * from file8_20 where doc_no = " & MyParn(xDoc_No.Caption)
con.Execute "Delete * from file8_20h where doc_no = " & MyParn(xDoc_No.Caption)
con.CommitTrans
myDel = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
Unload Me
End Sub

Private Sub xDate_Change()
Handlecontrols
End Sub

Private Sub xRest_Click()
xValue.Text = xRest.Caption
End Sub

Private Sub xRestCur_Click()
xValue.Text = xRestCur.Caption
End Sub
Private Sub xvalue_Change()
Handlecontrols
'xRest.Caption = Format(Val(xRest.Caption) - Val(xvalue.Text), "Fixed")
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub Handlecontrols()
CmdAdd.Enabled = IsDate(xDate.Text) And (Val(xRest.Caption) > 0 Or Val(xPaid.Caption) = 0) And Val(xValue.Text) > 0 And XBOX.BoundText <> ""
cmdEdit.Enabled = IsDate(xDate.Text) And Val(xValue.Text) > 0 And XBOX.BoundText <> "" And CmdDel.Enabled
End Sub
