VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form RepItem0 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
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
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   615
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
      Height          =   390
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   1515
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   390
      Left            =   675
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   75
      Width           =   1365
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
   End
End
Attribute VB_Name = "RepItem0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SourceTable1 As Recordset, SourceTable2 As Recordset, SourceTable3 As Recordset, sourceTable4 As Recordset, sourcetable5 As Recordset, sourcetable6 As Recordset
Dim TempTable As Recordset
Dim aReturn(2) As Double
Private Sub CmdApply_Click()
RepItem1
End Sub
Sub RepItem1()
tempdb.Execute "DELETE * FROM TEMP"
Dim TempTable As Recordset
Set TempTable = tempdb.OpenRecordset("Temp")

cString = "SELECT FILE7_30.DATE_INV, FILE7_30.ITEM, FILE7_30.boat, FILE7_30.inv_no, FILE1_10.DESCA, FILE7_30.REC, " & _
          "File1_10.DescA " & _
          " From file7_30 inner join file1_10 on file7_30.item = file1_10.item"
Set SourceTable1 = mydb.OpenRecordset(cString)

cString = "Select item,credit,Sum(Quant) as SumOfQuant" & _
          " From File6_20 group by Item,Credit  "
Set SourceTable2 = mydb.OpenRecordset(cString)

cString = "Select item,Credit,Sum(Quant) as SumOfQuant" & _
          " From File6_10 group by Item,Credit  "
Set SourceTable3 = mydb.OpenRecordset(cString)

cString = "Select item,Imp_No,Sum(Amount) as SumOfQuant" & _
          " From File1_80 group by Item,Imp_No  "
Set sourceTable4 = mydb.OpenRecordset(cString)

cString = "Select item,Imp_No,Sum(Amount) as SumOfQuant" & _
          " From File1_81 group by Item,Imp_No"
Set sourcetable5 = mydb.OpenRecordset(cString)

cString = "Select item,Imp_No,Sum(Amount) as SumOfQuant" & _
          " From File1_82 group by Item,Imp_No  "
Set sourcetable6 = mydb.OpenRecordset(cString)

Do While Not SourceTable1.EOF
    TempTable.AddNew
    RetQuant
    TempTable.str1 = SourceTable1.Item
    TempTable.str2 = SourceTable1.inv_No
    TempTable.str3 = SourceTable1.BOAT
    TempTable.str4 = SourceTable1.DESCA
    TempTable.val1 = SourceTable1.Rec

    TempTable.VAL3 = aReturn(0)

    TempTable.Date1 = SourceTable1.Date_INV
    TemptTable.str19 = firsttitle
    TempTable.str20 = Secondtitle

    TempTable.Update
    SourceTable1.MoveNext
Loop

Report1.ReportFileName = PublicPath & "\Reports\RepItm0.rpt"
Report1.DataFiles(0) = App.Path & "\Temp.mdb"
Report1.Action = 1
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command1_Click()
xDoc_No.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub RetQuant()
With SourceTable1
aReturn(0) = .Rec


SourceTable2.FindFirst "Item = " & MyParn(.Item) & _
                       " and Credit = " & MyParn(.inv_No)

SourceTable3.FindFirst "Item = " & MyParn(.Item) & _
                       " and Credit = " & MyParn(.inv_No)

sourceTable4.FindFirst "Item = " & MyParn(.Item) & _
                       " and Imp_No = " & MyParn(.inv_No)

sourcetable5.FindFirst "Item = " & MyParn(.Item) & _
                       " and Imp_No = " & MyParn(.inv_No)

sourcetable6.FindFirst "Item = " & MyParn(.Item) & _
                       " and Imp_No = " & MyParn(.inv_No)
If Not SourceTable2.NoMatch Then  'sales'
    aReturn(0) = aReturn(0) - SourceTable2!SumofQuant
   
End If

If Not SourceTable3.NoMatch Then 'retsales
    aReturn(0) = aReturn(0) + SourceTable3!SumofQuant
  
End If

If Not sourceTable4.NoMatch Then 'damage
    aReturn(0) = aReturn(0) - sourceTable4!SumofQuant

End If

If Not sourcetable5.NoMatch Then 'output
    aReturn(0) = aReturn(0) + sourcetable5!SumofQuant

End If

If Not sourcetable6.NoMatch Then 'input
    aReturn(0) = aReturn(0) - sourcetable6!SumofQuant

End If
End With
End Sub

