VERSION 5.00
Begin VB.Form NewItemFrm 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   13485
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "NewItemFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
With grid1
clist1 = StrList("Select code,desca from file1_50 order by desca")
cList2 = StrList("Select code,desca from file1_10sc order by desca")
cList3 = StrList("Select code,desca from FILE4_10 order by desca")
MyLoad
End With

End Sub
