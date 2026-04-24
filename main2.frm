VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form main2 
   BackColor       =   &H00808080&
   ClientHeight    =   8190
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   PaletteMode     =   2  'Custom
   RightToLeft     =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   11010
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport Report1 
      Left            =   375
      Top             =   150
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
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   4530
      Left            =   405
      Stretch         =   -1  'True
      Top             =   720
      Width           =   4590
   End
   Begin VB.Menu mnitems 
      Caption         =   "√’‰«›"
      NegotiatePosition=   1  'Left
      Begin VB.Menu XBDATA 
         Caption         =   "»Ì«‰«  «·«’‰«›"
      End
      Begin VB.Menu tmitemgroup 
         Caption         =   "„Ã„Ê⁄«  «·«’‰«›"
      End
      Begin VB.Menu tmitemgroupmain 
         Caption         =   "„Ã„Ê⁄«  «·«’‰«› «·—∆Ì”Ì…"
      End
      Begin VB.Menu tmsection 
         Caption         =   "«·«ﬁ”«„"
      End
      Begin VB.Menu XMSTORE 
         Caption         =   "„Œ«“‰"
      End
      Begin VB.Menu SEP1_1 
         Caption         =   "-"
      End
      Begin VB.Menu XMTRANS 
         Caption         =   " ÕÊÌ·«  »Ì‰ «·„Œ«“‰"
      End
      Begin VB.Menu tmDamage 
         Caption         =   "«· «·›"
      End
      Begin VB.Menu tminput 
         Caption         =   "Ê«—œ"
      End
      Begin VB.Menu SEP12 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu tmStock 
         Caption         =   "Ã—œ „Œ«“‰"
      End
      Begin VB.Menu SEP1_3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu xItemMove 
         Caption         =   "Õ—ﬂ… ’‰›"
      End
   End
   Begin VB.Menu mnclients 
      Caption         =   "⁄„·«¡"
      Begin VB.Menu xClientData 
         Caption         =   "»Ì«‰«  ⁄„·«¡"
      End
      Begin VB.Menu xClientGroup 
         Caption         =   "„Ã„Ê⁄«  ⁄„·«¡"
      End
      Begin VB.Menu LINE3 
         Caption         =   "-"
      End
      Begin VB.Menu xClientMove 
         Caption         =   "Õ—ﬂ… ⁄„·«¡"
      End
   End
   Begin VB.Menu mnVendorsx 
      Caption         =   "„Ê—œÌ‰"
      Begin VB.Menu xVendorData 
         Caption         =   "»Ì«‰«  „Ê—œÌ‰"
      End
      Begin VB.Menu xVendorGroup 
         Caption         =   "„Ã„Ê⁄«  „Ê—œÌ‰"
      End
      Begin VB.Menu LINE6 
         Caption         =   "-"
      End
      Begin VB.Menu xGrVEND 
         Caption         =   "Õ—ﬂ… „Ê—œÌ‰"
      End
   End
   Begin VB.Menu mnInvoice 
      Caption         =   "›Ê« Ì—"
      Begin VB.Menu xSales 
         Caption         =   "„»Ì⁄« "
      End
      Begin VB.Menu xRetSales 
         Caption         =   "„—œÊœ „»Ì⁄« "
      End
      Begin VB.Menu line4 
         Caption         =   "-"
      End
      Begin VB.Menu xpurchases 
         Caption         =   "„‘ —Ì« "
      End
      Begin VB.Menu XRETPURCH 
         Caption         =   "„—œÊœ „‘ —Ì« "
      End
      Begin VB.Menu XNHF 
         Caption         =   "-"
      End
      Begin VB.Menu tmimpcost 
         Caption         =   "›Ê« Ì—   ﬂ·›… «” Ì—«œÌ…"
      End
      Begin VB.Menu sp4_3 
         Caption         =   "-"
      End
      Begin VB.Menu XMBARPRINT 
         Caption         =   "ÿ»«⁄… «” Ìﬂ—“"
      End
      Begin VB.Menu tmbarcodeprint 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
      End
   End
   Begin VB.Menu mnMAINCash 
      Caption         =   "‰ﬁœÌ…"
      Begin VB.Menu xCashed 
         Caption         =   "„ﬁ»Ê÷«  „‰ ⁄„·«¡"
      End
      Begin VB.Menu xCASH2 
         Caption         =   "„œ›Ê⁄«  «·Ì «·„Ê—œÌ‰"
      End
      Begin VB.Menu sep5_1 
         Caption         =   "-"
      End
      Begin VB.Menu xmcash4 
         Caption         =   "„œ›Ê⁄«  «·Ì «·⁄„·«¡"
      End
      Begin VB.Menu xmcash3 
         Caption         =   "„ﬁ»Ê÷«  „‰ «·„Ê—œÌ‰"
      End
      Begin VB.Menu sep5_2 
         Caption         =   "-"
      End
      Begin VB.Menu tmcharge 
         Caption         =   "„’«—Ì›"
      End
      Begin VB.Menu tmchargecode 
         Caption         =   "«ﬂÊ«œ „’«—Ì›"
      End
      Begin VB.Menu tmchargemaincode 
         Caption         =   "«ﬂÊ«œ „’«—Ì› —∆Ì”Ì…"
      End
      Begin VB.Menu line5_3 
         Caption         =   "-"
      End
      Begin VB.Menu tmincome 
         Caption         =   " ”ÃÌ· «Ì—«œ« "
      End
      Begin VB.Menu tmincomecode 
         Caption         =   "«ﬂÊ«œ «Ì—«œ« "
      End
      Begin VB.Menu tmincomemaincode 
         Caption         =   "«ﬂÊ«œ «Ì—«œ«  —∆Ì”Ì…"
      End
      Begin VB.Menu Line5_4 
         Caption         =   "-"
      End
      Begin VB.Menu tmpart 
         Caption         =   "Ã«—Ì «·‘—ﬂ«¡"
      End
      Begin VB.Menu tmpart_code 
         Caption         =   "«ﬂÊ«œ «·‘—ﬂ«¡"
      End
      Begin VB.Menu line5_5 
         Caption         =   "-"
      End
      Begin VB.Menu tmbox 
         Caption         =   "«ﬂÊ«œ Œ“‰"
      End
      Begin VB.Menu tmboxtrans 
         Caption         =   " ÕÊÌ·«  Œ“‰"
      End
      Begin VB.Menu tmboxbal 
         Caption         =   "—’Ìœ «·Œ“‰…"
      End
   End
   Begin VB.Menu mnBank 
      Caption         =   "»‰Êﬂ"
      Begin VB.Menu tmBankData 
         Caption         =   "»Ì«‰«  «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankGrroup 
         Caption         =   "„Ã„Ê⁄«  «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankMove 
         Caption         =   "Õ—ﬂ… «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankItems 
         Caption         =   "«ﬂÊ«œ Õ—ﬂ… «·»‰ﬂ"
      End
      Begin VB.Menu tmBankInout 
         Caption         =   " ”ÃÌ· Õ—ﬂ… «·»‰Êﬂ"
      End
      Begin VB.Menu sep61 
         Caption         =   "-"
      End
      Begin VB.Menu tmChqIn 
         Caption         =   "√Ê—«ﬁ ﬁ»÷"
      End
      Begin VB.Menu tmChqOut 
         Caption         =   "√Ê—«ﬁ œ›⁄"
      End
      Begin VB.Menu sep62 
         Caption         =   "-"
      End
      Begin VB.Menu tmBankState 
         Caption         =   "ﬂ‘› Õ”«» »‰ﬂ"
      End
      Begin VB.Menu tmBankMoveTotal 
         Caption         =   "«Ã„«·Ì Õ—ﬂ… «·»‰ﬂ"
      End
      Begin VB.Menu tmbankDtl1 
         Caption         =   " ﬁ—Ì—  ›’Ì·Ì Õ—ﬂ… «·»‰ﬂ"
      End
   End
   Begin VB.Menu mnCASH 
      Caption         =   " ﬁ«—Ì—"
      Begin VB.Menu tm_vsitem 
         Caption         =   "≈Ã„«·Ï Õ—ﬂ… «·√’‰«› ··‘—ﬂ…"
      End
      Begin VB.Menu tmgroupsection 
         Caption         =   "≈Ã„«·Ì „Ã„Ê⁄«  Ê√ﬁ”«„"
      End
      Begin VB.Menu tmvsstore 
         Caption         =   "≈Ã„«·Ï —’Ìœ √’‰«› /  Ê“Ì⁄Â« ⁄·Ï «·„Œ«“‰"
      End
      Begin VB.Menu sep691 
         Caption         =   "-"
      End
      Begin VB.Menu ititemimport 
         Caption         =   "„Êﬁ›  ›’Ì·Ï √’‰«› —”«·… ≈” Ì—«œÌ…"
      End
      Begin VB.Menu TMCustImp 
         Caption         =   "≈Ã„«·Ï „»Ì⁄«  «·⁄„·«¡ „‰ —”«·… ≈” Ì—«œÌ…"
      End
      Begin VB.Menu sep661 
         Caption         =   "-"
      End
      Begin VB.Menu tmshopproft 
         Caption         =   "≈Ã„«·Ï „Êﬁ› › —… ··„Õ·"
      End
      Begin VB.Menu TMproftshop 
         Caption         =   " ›’Ì·Ï „Êﬁ› «’‰«›  «Õ„œ ”·Ì„«‰"
      End
      Begin VB.Menu sep64 
         Caption         =   "-"
      End
      Begin VB.Menu XTMBALSUPP 
         Caption         =   "≈Ã„«·Ï „»Ì⁄«  «’‰«› ··⁄„Ì· - „Ã„Ê⁄… ⁄„·«¡"
      End
      Begin VB.Menu tmTCUST 
         Caption         =   "≈Ã„«·Ï √—’œ… Ê  ⁄«„·«  «·⁄„·«¡"
      End
      Begin VB.Menu TMGRCUST 
         Caption         =   "≈Ã„«·Ï √—’œ… Ê  ⁄«„·«  „Ã„Ê⁄«  ⁄„·«¡"
      End
      Begin VB.Menu xfolcost 
         Caption         =   "„ «»⁄… ”⁄— «·‘—«¡"
         Visible         =   0   'False
      End
      Begin VB.Menu sep65 
         Caption         =   "-"
      End
      Begin VB.Menu storerep 
         Caption         =   " ﬁ«—Ì— „Œ«“‰"
      End
      Begin VB.Menu sep66 
         Caption         =   "-"
      End
      Begin VB.Menu tmClientReport 
         Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
      End
      Begin VB.Menu sep70 
         Caption         =   "-"
      End
      Begin VB.Menu m_SupRpt 
         Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
      End
      Begin VB.Menu sep69 
         Caption         =   "-"
      End
      Begin VB.Menu tmChargerep 
         Caption         =   " ﬁ«—Ì—  «·„’«—Ì›"
      End
      Begin VB.Menu sep67 
         Caption         =   "-"
      End
      Begin VB.Menu tmChqRep 
         Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
      End
      Begin VB.Menu sep68 
         Caption         =   "-"
      End
      Begin VB.Menu tmProftComp 
         Caption         =   "„Êﬁ› «·‘—ﬂ… „Õ·Ï Ê «” Ì—«œ"
      End
   End
   Begin VB.Menu mnServices 
      Caption         =   "Œœ„« "
      Begin VB.Menu tmsecurity 
         Caption         =   "’·«ÕÌ« "
      End
      Begin VB.Menu itChargImp 
         Caption         =   "√ﬂÊ«œ „’«—Ì› «·≈” Ì—«œÌ…"
      End
      Begin VB.Menu ItPostItem 
         Caption         =   " ÕœÌÀ √’‰«› «·„Õ·"
      End
      Begin VB.Menu xMSal 
         Caption         =   "√ﬂÊ«œ »«∆⁄Ì‰"
      End
      Begin VB.Menu line_exit 
         Caption         =   "-"
      End
      Begin VB.Menu xExit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "Main2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tSecurity As Recordset
Private Sub Form_Load()
SetKbLayout Lang_AR
If Not bSupermode Then LoadMenu
rdItem.Open "FILE1_10", con, adOpenStatic, adLockReadOnly
Firsttitle = "Ali's Market "
tmshopproft.Visible = bopt1
tmProftComp.Visible = bopt1
End Sub

Private Sub ItemMove_Click()
Load itemMove
itemMove.Show 1
End Sub
Private Sub StoreMove_Click()
Load StoreMove
StoreMove.Show 1
End Sub
Private Sub Stores_Click()
flag.myFlag = 1
flag.Show 1
End Sub
Private Sub report1_Click()
Load Form4
Form4.Show 1
End Sub
Private Sub bal_box_Click()
    Load BalBox
    BalBox.Show 1
End Sub
Private Sub BANK_MOVE_Click()
Vs_Bank.Show 1
End Sub
Private Sub CHQREP_Click()
Load repchq
repchq.Show 1
End Sub
Private Sub m_ChargeMain_Click()
flag.myFlag = 7
flag.Show 1
End Sub
Private Sub m_debtCode_Click()
publicFlag = 1
Load AssetCode
AssetCode.Show 1
End Sub
Private Sub m_DebtGrp_Click()
publicFlag = 15
Load flag
flag.Show 1
End Sub
Private Sub m_Income_Click()
publicFlag = 1
Load ChargeSub
ChargeSub.Show 1
End Sub

Private Sub m_IncomeMain_Click()
publicFlag = 13
Load flag
flag.Show 1
End Sub
Private Sub M_SALMED_Click()
Load SalMed
SalMed.Show 1
End Sub

Private Sub itChargImp_Click()
ReDim aPublic(5)
aPublic(0) = "FILE7_60CH_CODE"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„’—Ê›"
aPublic(4) = "»Ì«‰ "
aPublic(5) = "√ﬂÊ«œ „’«—Ì› ≈” Ì—«œÌ…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1

End Sub

Private Sub ititemimport_Click()
    VsImpItem.Show 1
End Sub

Private Sub ItPostItem_Click()
    CopyItem.Show 1
End Sub
Private Sub m_SupRpt_Click()
    rpSup.Show 1
End Sub
Private Sub m_daysales_Click()
    Load DaySale
    DaySale.Show 1
End Sub
Private Sub mChargerep_Click()
    RepCharge.Show 1
End Sub
Private Sub MCURR_Click()
publicFlag = 6
Load flag
flag.Show 1
End Sub
Private Sub mInput_Click()
publicFlag = 0
Load Vs_Input
Vs_Input.Show 1
End Sub

Private Sub MnTSalMan_Click()
vsman.Show 1
End Sub

Private Sub Mony_Box_Click()
publicFlag = 5
flag.Show 1
End Sub
Private Sub mPricelist_Click()
Load PriceList
PriceList.Show 1
End Sub
Private Sub mTransAction_Click()
Load Post
Post.Show 1
End Sub
Private Sub mPrevInv_Click()
    Load PrevInv
    PrevInv.Show 1
End Sub

Private Sub mxclose_Click()
FrmClose.Show 1
End Sub

Private Sub mXsec_Click()
    Security.Show 1
End Sub

Private Sub mxtrans_Click()
    TransBox.Show 1
End Sub
Private Sub PHO_MAI_Click()
Load phone
phone.Show 1
End Sub
Private Sub RBAN_Click()
    Load RepBank
    RepBank.Show 1
End Sub
Private Sub utillll_Click()
Load Utility2
Utility2.Show 1
End Sub
Private Sub Trans_Box_Click()
TransBox.Show 1
End Sub

Private Sub tm_vsitem_Click()
    grditem1.Show 1
End Sub

Private Sub tmBankData_Click()
bankfrm.bEdit = True
bankfrm.Show 1
End Sub
Private Sub tmbankDtl1_Click()
    rpBank3.Show 1
End Sub
Private Sub tmBankGrroup_Click()
ReDim aPublic(5)
aPublic(0) = "FILE5_50"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
aPublic(4) = "≈”„ «·„Ã„Ê⁄…"
aPublic(5) = "„Ã„Ê⁄«  «·»‰Êﬂ"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmBankInout_Click()
'    bEdit = RetSec(tmItem.Name)
'    bDel = RetSec(tmItem.Name, "Del")
    bankinoutfrm.bEdit = True
    bankinoutfrm.Show 1
End Sub
Private Sub tmBankItems_Click()
ReDim aPublic(5)
aPublic(0) = "FILE5_00"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·Õ—ﬂ…"
aPublic(4) = "≈”„ «·Õ—ﬂ…"
aPublic(5) = "«ﬂÊ«œ Õ—ﬂ… «·»‰Êﬂ"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmBankMove_Click()
    BankMovefrm.Show 1
End Sub
Private Sub tmBankMoveTotal_Click()
    rpBank2.Show 1
End Sub
Private Sub tmBankState_Click()
    bankStatefrm.Show 1
End Sub

Private Sub tmbarcodeprint_Click()
    Dream_Bar.Show 1
End Sub

Private Sub tmbox_Click()
    Boxfrm.Show 1
End Sub
Private Sub tmboxbal_Click()
    BalBox.Show 1
End Sub
Private Sub tmboxtrans_Click()
    bEdit = True
    boxtransfrm.Show 1
End Sub
Private Sub tmcharge_Click()
chargefrm.myPublic = 1
chargefrm.Show 1
End Sub
Private Sub tmchargecode_Click()
chargecodefrm.bEdit = True
chargecodefrm.myPublic = 1
chargecodefrm.Show 1
End Sub
Private Sub tmchargemaincode_Click()
ReDim aPublic(6)
aPublic(0) = "FILE8_52"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„’—Ê›"
aPublic(4) = "»Ì«‰ «·„’—Ê›"
aPublic(5) = "„’«—Ì› —∆Ì”Ì…"
aPublic(6) = 3
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub
Private Sub tmChargerep_Click()
    rpCharge.Show 1
End Sub
Private Sub tmChqIn_Click()
'bEdit = RetSec(tmCollectPaper.Name)
    publicFlag = 1
    bEdit = True
    chq.Show 1
End Sub
Private Sub tmChqOut_Click()
    publicFlag = 2
    chq.Show 1
End Sub
Private Sub tmChqRep_Click()
    rpChq.Show 1
End Sub
Private Sub tmClientReport_Click()
    rpClient.Show 1
End Sub

Private Sub TMCustImp_Click()
    CustSalesImp.Show 1
End Sub

Private Sub tmDamage_Click()
    damagefrm.Show 1
End Sub

Private Sub TMGRCUST_Click()
    VsTBalCustGR.Show 1
End Sub
Private Sub tmgroupsection_Click()
    VsTGroup.Show 1
End Sub
Private Sub tmimpcost_Click()
    impcostfrm.bEdit = True
    impcostfrm.Show 1
End Sub
Private Sub tmincome_Click()
    chargefrm.myPublic = 2
    chargefrm.Show 1
End Sub
Private Sub tmincomecode_Click()
    chargecodefrm.bEdit = True
    chargecodefrm.myPublic = 2
    chargecodefrm.Show 1
End Sub
Private Sub tmincomemaincode_Click()
ReDim aPublic(5)
aPublic(0) = "FILE8_61"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·«Ì—«œ"
aPublic(4) = "»Ì«‰ «·«Ì—«œ"
aPublic(5) = "«Ì—«œ«  —∆Ì”Ì…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmindust_Click()
industFrm.Show 1
End Sub
Private Sub tminput_Click()
    inputfrm.Show 1
End Sub
Private Sub tmitemgroup_Click()
itemsGroupFrm.bEdit = True
itemsGroupFrm.Show 1
End Sub
Private Sub tmitemgroupmain_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub

Private Sub tmpart_Click()
bEdit = True
partfrm.Show 1
End Sub

Private Sub tmpart_code_Click()
ReDim aPublic(5)
aPublic(0) = "FILE8_71"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·‘—Ìﬂ"
aPublic(4) = "≈”„ «·‘—Ìﬂ"
aPublic(5) = "«ﬂÊ«œ «·‘—ﬂ«¡"
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub

Private Sub tmProftComp_Click()
    Tproft_Comp.Show 1
End Sub

Private Sub TMproftshop_Click()
    VsTProftShop.Show 1
End Sub

Private Sub tmsection_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_10SC"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "√ﬁ”«„ «·«’‰«›"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmsecurity_Click()
Security.Show 1
End Sub

Private Sub tmshopproft_Click()
    Tproft_shop.Show 1
End Sub

Private Sub tmStock_Click()
bEdit = True
StockFrm.Show 1
End Sub
Private Sub tt_cust_Click()
    VsC.Show 1
    End Sub
Private Sub tt_item_Click()
     grditem1.Show 1
End Sub
Private Sub tt_vend_Click()
    VsSupp.Show 1
End Sub
Private Sub x_WAG_Click()
    Load Emp
    Emp.Show 1
End Sub
Private Sub XAbount_Click()
    Load About
    About.Show 1
End Sub
Private Sub xCASH1_Click()
    publicFlag = 4
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub

Private Sub tmTCUST_Click()
    VsTBalCust.Show 1
End Sub
Private Sub tmvsstore_Click()
    VsTStore.Show 1
End Sub
Private Sub xCASH2_Click()
bEdit = True
Cashfrm.myPublic = 2
Cashfrm.Show 1
End Sub
Private Sub xCen_Cost_Click()
publicFlag = 5
Load flag
flag.Show 1
End Sub
Private Sub xprofff_Click()
    Load Profit
    Profit.Show 1
End Sub

Private Sub xchksupp_Click()
    chq2.Show 1
End Sub
Private Sub xfolcost_Click()
    VsPrice2.Show 1
End Sub
Private Sub XFOLPRICE_Click()
    VsPrice.Show 1
End Sub

Private Sub xGrVEND_Click()
supMovefrm.Show 1
End Sub

Private Sub xmacc_Click()
    FrmProft.Show 1
End Sub
Private Sub XMBALBOX_Click()
    BalBox.Show 1
End Sub
Private Sub XMBALSTORE_Click()
    StoreBal.Show 1
End Sub
Private Sub XMBARPRINT_Click()
     Morsh_Bar.Show 1
End Sub
Private Sub xmcash3_Click()
bEdit = True
Cashfrm.myPublic = 4
Cashfrm.Show 1
End Sub
Private Sub xmcash4_Click()
bEdit = True
Cashfrm.myPublic = 3
Cashfrm.Show 1
End Sub
Private Sub XMCODEBOX_Click()
    MonyBox.Show 1
End Sub
Private Sub XMPAY_Click()
    Vs_Pay.Show 1
End Sub

Private Sub XMREDEM_Click()
    VsSal.Show 1
End Sub
Private Sub xMReSTORE_Click()
VsSal.Show 1
End Sub
Private Sub XMREDEM2_Click()
    Redem.Show 1
End Sub

Private Sub xMSal_Click()
ReDim aLocal(6)
aLocal(0) = "FILE6_25"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ «·»«∆⁄"
aLocal(4) = "≈”„ «·»«∆⁄"
aLocal(5) = "«ﬂÊ«œ »«∆⁄Ì‰"
aLocal(6) = 1
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub XMSALCODE_Click()
Vstsalsupp.Show 1
End Sub
Private Sub XMSTORE_Click()
ReDim aLocal(6)
aLocal(0) = "FILE0_40"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ «·„Œ“‰"
aLocal(4) = "»Ì«‰ «·„Œ“‰"
aLocal(5) = " ”ÃÌ· « ·„Œ«“‰"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub XMTCUSTSAL_Click()
    F_SalCust.Show 1
End Sub
Private Sub XMTRANS_Click()
    Transfrm.Show 1
End Sub
Private Sub XMTSTORE_Click()
    vsstore.Show 1
End Sub
Private Sub XMVISA_Click()
    SalVisa.Show 1
End Sub
Private Sub XRETPURCH_Click()
    Purchasefrm.myPublic = 1
    Purchasefrm.Show 1
End Sub
Private Sub storerep_Click()
    rpItem.Show 1
End Sub
Private Sub XBDATA_Click()
'itemsgrdFrm.bEdit = RetSec(XBDATA.Name)
'itemsgrdFrm.Show 1
itemsfrm.Show 1
End Sub
Private Sub xCASH3_Click()
    publicFlag = 3
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub
Private Sub xCASH4_Click()
    publicFlag = 2
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub
Private Sub xCashed_Click()
bEdit = True
Cashfrm1.Show 1
End Sub
Private Sub XcHARG_Click()
    publicFlag = 1
    Load Vs_Charg
    Vs_Charg.Show 1
End Sub
Private Sub XcHARGE_Click()
    publicFlag = 7
    Load flag
    flag.Show 1
End Sub
Private Sub XCHQ_Click()
    publicFlag = 1
    chq.Show 1
End Sub
Private Sub xClientData_Click()
Clients.myFlag = 1
Clients.Show 1
End Sub
Private Sub xClientGroup_Click()
ReDim aLocal(6)
aLocal(0) = "FILE3_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ"
aLocal(4) = "≈”„ «·„Ã„Ê⁄…"
aLocal(5) = " ”ÃÌ· «·⁄„·«¡"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub xClientMove_Click()
    ClientMoveFrm.Show 1
End Sub
Private Sub xClientReport_Click()
Load ClientReports
ClientReports.Show 1
End Sub
Private Sub xCompGroup_Click()
End Sub
Private Sub xCredit_Click()
    Load CREDIT
    CREDIT.Show 1
End Sub
Private Sub XCREDITREP_Click()
    Load RepCredit
    RepCredit.Show 1
End Sub
Private Sub XDATA_Click()
Load Bank
Bank.Show 1
End Sub
Private Sub xDel_Move_Click()
If MsgBox("«·€«¡ﬂ· «·»Ì«‰«   : Â· «‰  „Ê«›ﬁ ø", 4 + 256) <> 6 Then Exit Sub
Dim aString(12)
aString(0) = "file0_10"
aString(1) = "file1_10"
aString(2) = "file1_11"
aString(3) = "file1_30"
aString(4) = "file1_60"
aString(5) = "file1_70"
aString(6) = "file1_50"
aString(7) = "file3_10"
aString(8) = "file4_10"
aString(9) = "file6_10"
aString(10) = "file6_20"
aString(11) = "file7_20"
For i = 0 To UBound(aString) - 1
    cString = "Delete Distinctrow " & aString(i) & ".*" & " From " & aString(i)
    mydb.Execute cString
Next
End Sub
Private Sub xExit_Click()
    End
End Sub
Private Sub xGroup_Click()
Load ItemsGrp
ItemsGrp.Show 1
End Sub
Private Sub xIMPORT_Click()
Load Import
Import.Show 1
End Sub
Private Sub xinventory_Click()
'Load inventory
'inventory.Show 1
Vs_Stock.Show 1
End Sub
Private Sub xItemMove_Click()
StoreMove.Show 1
End Sub
Private Sub xMainGroup_Click()
publicFlag = 2
Load flag
flag.Show 1
End Sub
Private Sub xMortal_Click()
publicFlag = 2
Load Vs_Input
Vs_Input.Show 1
End Sub
Private Sub xOutPut_Click()
publicFlag = 1
Load Vs_Input
Vs_Input.Show 1
End Sub
Private Sub xpurchases_Click()
Purchasefrm.myPublic = 0
Purchasefrm.Show 1
End Sub
Private Sub xRepSAlMan_Click()
    Load rsalman
    rsalman.Show 1
End Sub
Private Sub xRetSales_Click()
salesfrm.myPublic = 1
salesfrm.Show 1
End Sub
Private Sub xSales_Click()
'publicFlag = 1
'Vs_Inv.Show 1
salesfrm.myPublic = 0
salesfrm.Show 1
'publicFlag = 0
'Load Invoice
'Invoice.Show 1
End Sub
Private Sub xStoreMove_Click()
Load StoreMove
StoreMove.Show 1
End Sub
Private Sub xStores_Click()
publicFlag = 1
Load flag
flag.Show 1
End Sub
Private Sub xStoreTrans_Click()
    Load Vs_Trans
    Vs_Trans.Show 1
End Sub
Private Sub xutil11_Click()
Load Utility
Utility.Show 1
End Sub

Private Sub XTMBALSUPP_Click()
    VsTCustSales.Show 1
End Sub
Private Sub xVendorData_Click()
Clients.myFlag = 2
Clients.Show 1
End Sub
Private Sub xVendorGroup_Click()
ReDim aLocal(6)
aLocal(0) = "FILE4_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ"
aLocal(4) = "≈”„ «·„Ã„Ê⁄…"
aLocal(5) = " ”ÃÌ· «·„Ê—œ"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub Form_Resize()
Fiximage App.Path & "\graph\01-02.jpg"
'Fiximage "C:\WINDOWS\DESKTOP\TEST\01.JPG"
End Sub
Private Sub Fiximage(cImage)
Image1.Top = 0
Image1.Left = 0
'Image1.Picture = LoadPicture(App.Path & "\graph\main.jpg")
2001

Image1.Width = Me.Width
If Me.Height > 500 Then Image1.Height = Me.Height - 500
End Sub

Private Sub xxComp_Click()
CompData.Show 1
End Sub
Private Sub LoadMenu()
For i = 0 To main.Count - 1
    If TypeOf main(i) Is Menu And Mid(main(i).Name, 1, 2) = "mn" Then
        main(i).Visible = False
    End If
Next

Err.Clear
On Error Resume Next
For i = 0 To main.Count - 1
    If TypeOf main(i) Is Menu And Mid(main(i).Name, 1, 2) <> "mn" Then
        sectable.Find "control = " & MyParn(main(i).Name), , adSearchForward, adBookmarkFirst
        If Not sectable.EOF Then
            If sectable!Visible Then
                main(sectable!MainMenu).Visible = True
                main(i).Visible = True
            Else
                main(i).Visible = False
            End If
        End If
    End If
Next
End Sub
