VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Tproft_Comp 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2910
   ScaleWidth      =   4920
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   345
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   3480
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ãä äÇÑíÎ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì ÊÇÑíÎ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   345
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1725
      Width           =   3480
      Begin VB.CommandButton Cmd_Exit 
         Caption         =   "ÎÜÜÜÑæÌ"
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1245
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ØÈÇÚÉ ÇáãæÞÝ"
         Height          =   465
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1800
      End
   End
End
Attribute VB_Name = "Tproft_Comp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public ConShop As New ADODB.Connection
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub Cmd_Print_Click()
Dim temptable As New ADODB.Recordset, aHeader(1)
Dim n1 As Double, n2 As Double, n3 As Double, n31 As Double, n4 As Double, n6 As Double
Dim n61 As Double, n7 As Double, n13 As Double, n14 As Double, n15 As Double, n16 As Double
Dim n17 As Double, n12 As Double
Dim SalTable As New ADODB.Recordset
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If Not (IsDate(xdate1.Text) And IsDate(xDate2.Text)) Then
    MsgBox "ÇáÊÇÑíÎ ÛíÑ ÕÍíÍ"
    Exit Sub
End If
    cWhere = " date >= " & DateSq(xdate1.Text) & " AND DATE <= " & DateSq(xDate2.Text)

    temptable.AddNew
    n1 = Val(GetDesca("SELECT Sum(([IN]-[OUT])* FILE1_10.[COST] )  FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  Where FILE1_10.IMPORT = 1 AND  DATE < " & DateSq(xdate1.Text)) & "")
    n2 = Val(GetDesca("SELECT Sum(([IN]-[OUT])* FILE1_10.[COST] )  FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  Where FILE1_10.IMPORT = 0  AND  DATE < " & DateSq(xdate1.Text)) & "")
    temptable!str1 = "ÊÞíã ÃÑÕÏÉ ÇáÇÕäÇÝ áÃæá ÇáÝÊÑÉ "
    temptable!val1 = n1
    temptable!val2 = n2
    
    temptable!Val10 = 1
    temptable!val9 = 1
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update

    n1 = Val(GetDesca("select sum((FILE1_11.[IN] - FILE1_11.[OUT])* FILE1_11.PRICE ) as tall FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM where FILE1_10.IMPORT = 1 AND (FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7' ) AND " & cWhere) & "")
    n2 = Val(GetDesca("select sum((FILE1_11.[IN] - FILE1_11.[OUT])* FILE1_11.PRICE ) as tall FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM where FILE1_10.IMPORT = 0 AND (FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7' ) AND " & cWhere) & "")
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ãÔÊÑíÇÊ ÎáÇá ÇáÝÊÑÉ "
    temptable!val1 = n1
    temptable!val2 = n2
    temptable!val9 = 2
    temptable!Val10 = 2
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update

    cStr1 = " SELECT FILE1_10.IMPORT, FILE3_10.cust, Sum([TOTAL]*(1-[RATEDISC])) AS TSALES , Sum(([OUT]-[IN])*[FILE1_10].[COST]) AS TCOST FROM FILE1_10 INNER JOIN (FILE3_10 INNER JOIN NET_ALLSALES ON FILE3_10.CODE = NET_ALLSALES.code) ON FILE1_10.ITEM = NET_ALLSALES.ITEM WHERE " & cWhere & " GROUP BY FILE1_10.IMPORT, FILE3_10.cust"
    SalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    
    SalTable.Filter = " IMPORT = 1 AND CUST = 0 "
    If Not SalTable.EOF Then
        n1 = Val(SalTable!tSALES & "")
        n2 = Val(SalTable!Tcost & "")
    End If
    
    SalTable.Filter = " IMPORT = 0 AND CUST = 0 "
    If Not SalTable.EOF Then
        n3 = Val(SalTable!tSALES & "")
        n4 = Val(SalTable!Tcost & "")
    End If
    
    SalTable.Filter = " IMPORT = 1 AND CUST = 1 "
    If Not SalTable.EOF Then
        n5 = Val(SalTable!tSALES & "")
        n6 = Val(SalTable!Tcost & "")
    End If
    
    SalTable.Filter = " IMPORT = 0 AND CUST = 1 "
    If Not SalTable.EOF Then
        n7 = Val(SalTable!tSALES & "")
        n8 = Val(SalTable!Tcost & "")
    End If
    
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÞíãÉ ãÈíÚÇÊ ÌãáÉ ááÝÊÑÉ "
    temptable!val1 = n1
    temptable!val2 = n3
    temptable!Val10 = 3
    temptable!val9 = 3
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÊßáÝÉ ãÈíÚÇÊ ÌãáÉ ááÝÊÑÉ "
    temptable!val1 = n2
    temptable!val2 = n4
    temptable!Val10 = 4
    temptable!val9 = 3
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÑÈÍ ãÈíÚÇÊ ÌãáÉ ááÝÊÑÉ "
    temptable!val1 = n1 - n2
    temptable!val2 = n3 - n4
    temptable!Val10 = 5
    temptable!val9 = 3
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÞíãÉ ãÈíÚÇÊ ÞØÇÚì áÝÊÑÉ "
    temptable!val1 = n5
    temptable!val2 = n7
    temptable!val9 = 4
    temptable!Val10 = 7
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÊßáÝÉ ãÈíÚÇÊ ÞØÇÚì ááÝÊÑÉ "
    temptable!val1 = n6
    temptable!val2 = n8
    temptable!Val10 = 8
    temptable!val9 = 4
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÑÈÍ ãÈíÚÇÊ ÞØÇÚì ááÝÊÑÉ "
    temptable!val1 = n5 - n6
    temptable!val2 = n7 - n8
    temptable!val9 = 4
    temptable!Val10 = 9
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
    
    temptable.AddNew
    temptable!str1 = "ÕÇÝì ÞíãÉ ÇáãÈíÚÇÊ ááÝÊÑÉ "
    temptable!val1 = n1 + n5
    temptable!val2 = n3 + n7
    temptable!Val10 = 10
    temptable!val9 = 5
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÕÇÝì ÊßáÝÉ ÇáãÈíÚÇÊ ááÝÊÑÉ "
    temptable!val1 = n2 + n6
    temptable!val2 = n4 + n8
    temptable!Val10 = 11
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable!val9 = 5
    temptable.Update
        
    temptable.AddNew
    temptable!str1 = "ÅÌãÇáì ÑÈÍ ÇáãÈíÚÇÊ ááÝÊÑÉ "
    temptable!val1 = n1 - n2 + (n5 - n6)
    temptable!val2 = n3 - n4 + (n7 - n8)
    temptable!val9 = 5
    temptable!Val10 = 12
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update
        
        
    n1 = Val(GetDesca("SELECT Sum(([IN]-[OUT])* FILE1_10.[COST] )  FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  Where FILE1_10.IMPORT = 1 AND  DATE <=" & DateSq(xDate2.Text)) & "")
    n2 = Val(GetDesca("SELECT Sum(([IN]-[OUT])* FILE1_10.[COST] )  FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  Where FILE1_10.IMPORT = 0 AND  DATE <=" & DateSq(xDate2.Text)) & "")
    temptable.AddNew
    temptable!str1 = "ÊÞíã ÃÑÕÏÉ ÇáÇÕäÇÝ áÊÇÑíÎ " & xDate2.Text
    temptable!val1 = n1
    temptable!val2 = n2
    temptable!Val10 = 13
    temptable!val9 = 6
    temptable!str21 = "ãä ÊÇÑíÎ " & xdate1.Text & " ÍÊì ÊÇÑíÎ " & xDate2.Text
    temptable.Update

contemp.BeginTrans
contemp.CommitTrans


main.Report1.ReportFileName = App.Path & "\Reports\T_PROFTCOMP.RPT"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1

temptable.Close
Set temptable = Nothing
End Sub
Function GetDescaSHOP(pString) As String
Dim loctable As New ADODB.Recordset
loctable.Open pString, ConShop, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.BOF And loctable.EOF) Then GetDescaSHOP = loctable(0) & ""
loctable.Close
Set loctable = Nothing
End Function
Private Sub Form_Load()
OpenCon con
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
