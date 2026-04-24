VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Memrep20 
   Caption         =   "»Ì«‰«  «·«⁄÷«¡"
   ClientHeight    =   3570
   ClientLeft      =   1065
   ClientTop       =   1875
   ClientWidth     =   10080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   10080
   Begin VB.CommandButton cmdExit 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   2295
      Width           =   1500
   End
   Begin VB.CommandButton cmdClear 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1665
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "„”Õ «·ﬂ·"
      Top             =   2295
      Width           =   1500
   End
   Begin VB.CommandButton CmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3195
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   2295
      Width           =   1500
   End
   Begin VB.Frame Frame3 
      Height          =   1275
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   990
      Width           =   2760
      Begin VB.CheckBox Check4 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ Õ«›ŸÌ «·⁄÷ÊÌ… ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   810
         Width           =   2400
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «·„ Ê›ÌÌ‰ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   135
         Width           =   1950
      End
      Begin VB.CheckBox Check2 
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ ”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   495
         Width           =   2400
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   990
      Width           =   2400
      Begin VB.CheckBox Check5 
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ Õ«›Ÿ «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   810
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ ”«ﬁÿÌ «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   495
         Width           =   2040
      End
      Begin VB.CheckBox xDied 
         Appearance      =   0  'Flat
         Caption         =   "»œÊ‰ «·„ Ê›ÌÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   5355
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   4515
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Tag             =   "D"
         Top             =   540
         Width           =   1410
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Tag             =   "D"
         Top             =   540
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xType 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   900
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1305
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdYear 
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   1710
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   661
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Œ «— «·„Ê”„"
         ButtonStyle     =   3
      End
      Begin VB.Label Label2 
         Caption         =   "·„ Ì”œœ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "«Œ— ”œ«œ ›Ì"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1350
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… «·⁄÷ÊÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "„”œœ „‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "”œœ „Ê”„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   960
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3510
      Top             =   225
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "Memrep20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdOk_Click()
Select Case publicFlag
Case 20
    RunReport
Case 21
    RunReport21
End Select
End Sub
Private Function doPrint()
Dim temptable As New ADODB.Recordset, SourceTable As New ADODB.Recordset
Dim aHeader(4)
cString = "Select Count(*) as countofMember,file1_10.[type],TYPE_CODES.DescA From File1_10 inner join TYPE_codes on File1_10.TYPE = TYPE_codes.Codes"


If IsNumeric(cmdYear(0).Tag) Then
    aHeader(1) = "«·–Ì‰ ”œœÊ« „Ê”„ " & cmdYear(0).Caption
    cWhere = cWhere & turn(cWhere, " and ") & "FILE1_10.CODE IN(SELECT CODE FROM FILE6_20H WHERE YEAR"
End If

If xDied.Value > 0 Then
    aHeader(2) = "»œÊ‰ «·„ Ê›ÌÌ‰"
    cWhere = cWhere & turn(cWhere, " and ") & " (Died = false)"
End If

If xDiedOnly.Value = 1 Then
    aHeader(3) = "«·„ Ê›ÌÌ‰ ›ﬁÿ"
    cWhere = cWhere & turn(cWhere, " and ") & " Died"
End If

If xSplit.Value = 1 Then
    aHeader(4) = "«·”«ﬁÿÌ «·⁄÷ÊÌ… ›ﬁÿ"
    cWhere = cWhere & turn(cWhere) & " (drop)"
End If
cString = cString & turn(cWhere, " where ") & cWhere
cString = cString & " Group By File1_10.[Section],section_codes.Desca"

SourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
contemp.Execute "delete * from temp"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adcmtext

With SourceTable
Do Until SourceTable.EOF
    temptable.AddNew
    temptable!val10 = SourceTable![Section]
    temptable!val1 = SourceTable!CountOfMember
    temptable!val2 = 0
    temptable!val3 = 0
    temptable!val4 = 0
    temptable!str2 = SourceTable!desca
    temptable!str11 = TurnValue(aHeader(0))
    temptable!str12 = TurnValue(retHeader(aHeader, 1, 4))
    temptable.Update
    SourceTable.MoveNext
Loop

temptable.Requery
If xAppend.Value > 0 Then
    cField1 = myiif("Relation = 1", 1) & " as countofwife"
    cField2 = myiif("Relation = 2", 1) & " as countofSon"
    cField3 = myiif("Relation > 2", 1) & " as countofRel"
    
    cString = "Select File1_10.[Section],section_codes.Desca," & _
              cField1 & "," & cField2 & "," & cField3 & _
              " From (File1_10 inner join file1_11 on file1_10.code = file1_11.Member) inner join section_codes on file1_10.[section] = section_codes.code "
    cString = cString & turn(cWhere, " where ") & cWhere
    cString = cString & " Group by File1_10.[Section] ,section_codes.Desca"
    
    SourceTable.Close
    SourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adcmtext
    Do Until SourceTable.EOF
            ' temptable.edit
        temptable.AddNew
        temptable!str2 = SourceTable!desca
        temptable!val10 = SourceTable![Section]
        temptable!val2 = SourceTable!countOfWife
        temptable!val3 = SourceTable!countOfSon
        temptable!val4 = SourceTable!countOfRel
        temptable!str11 = TurnValue(aHeader(0))
        temptable!str12 = TurnValue(retHeader(aHeader, 1, 4))
        temptable.Update
        SourceTable.MoveNext
    Loop
End If


If temptable.BOF And temptable.EOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
Else
    con.BeginTrans
    con.CommitTrans
    main.Report1.ReportFileName = MainPath & "\rpt\mRep21.rpt"
    main.Report1.DataFiles(0) = cTempPath
    main.Report1.Action = 1
End If
Set temptable = Nothing
Set SourceTable = Nothing
End With
End Function
Private Sub Form_Load()
End Sub

Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
