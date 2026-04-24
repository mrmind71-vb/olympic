VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "œ« « »—Ê"
   ClientHeight    =   9990
   ClientLeft      =   165
   ClientTop       =   -585
   ClientWidth     =   15120
   LinkTopic       =   "MDIForm1"
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport REPORT1 
      Left            =   1935
      Top             =   495
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
   Begin VB.Menu mn_items 
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…"
      NegotiatePosition=   1  'Left
      Begin VB.Menu tm_member 
         Caption         =   "»Ì«‰«  «·«⁄÷«¡"
      End
      Begin VB.Menu tm_paid 
         Caption         =   "„ÿ«·»«  «·«⁄÷«¡ «·⁄«„·Ì‰"
      End
      Begin VB.Menu tm_card1 
         Caption         =   "ÿ»«⁄… ﬂ«—‰ÌÂ«  «·«⁄÷«¡"
      End
      Begin VB.Menu tm_cardqry_chair 
         Caption         =   "ÿ»«⁄… ﬂ«—‰ÌÂ«  «⁄÷«¡ „Ã·” «·«œ«—…"
      End
      Begin VB.Menu tm_close_day 
         Caption         =   "”œ«œ «·Œ“‰…"
      End
      Begin VB.Menu sp_1_1 
         Caption         =   "-"
      End
      Begin VB.Menu tm_member_i 
         Caption         =   "»Ì«‰«  «⁄÷«¡ «· ﬁ”Ìÿ"
      End
      Begin VB.Menu tm_paid_install 
         Caption         =   "„ÿ«·»«  «·«⁄÷«¡ «· ﬁ”Ìÿ"
      End
      Begin VB.Menu tm_card_inv 
         Caption         =   "ÿ»«⁄… ﬂ«—‰ÌÂ«  «·«⁄÷«¡ «·„ﬁ”ÿ…"
      End
      Begin VB.Menu sp_1_4 
         Caption         =   "-"
      End
      Begin VB.Menu tm_member_inv 
         Caption         =   "⁄÷ÊÌ… œ⁄Ê…"
      End
      Begin VB.Menu tm_print_card_h 
         Caption         =   "ÿ»«⁄… ﬂ«—‰ÌÂ«  «·⁄÷ÊÌ… «·œ⁄Ê…"
      End
      Begin VB.Menu sp_1_5 
         Caption         =   "-"
      End
      Begin VB.Menu tm_items 
         Caption         =   "»‰Êœ ”œ«œ «·«⁄÷«¡"
      End
      Begin VB.Menu tm_worker 
         Caption         =   "»Ì«‰«  „ÊŸ›Ì‰"
      End
      Begin VB.Menu tm_card_worker 
         Caption         =   "ÿ»«⁄«  ﬂ«—‰ÌÂ«  «·⁄«„·Ì‰"
      End
      Begin VB.Menu sp_1_6 
         Caption         =   "-"
      End
      Begin VB.Menu tm_realtion_codes 
         Caption         =   "«ﬂÊ«œ «·ﬁ—«»…"
      End
   End
   Begin VB.Menu mn_sport 
      Caption         =   "«·‰‘«ÿ «·—Ì«÷Ì"
      Begin VB.Menu tm_members 
         Caption         =   "»Ì«‰«  «·«⁄÷«¡ ··‰‘«ÿ «·—Ì«÷Ì"
      End
      Begin VB.Menu tm_exit_sport 
         Caption         =   "Œ—ÊÃ"
      End
   End
   Begin VB.Menu mn_meeting 
      Caption         =   "«·Ã„⁄Ì… «·⁄„Ê„Ì…"
      Begin VB.Menu tm_meeting 
         Caption         =   "»Ì«‰«  «·Ã„⁄Ì… «·⁄„Ê„Ì…"
      End
      Begin VB.Menu tm_meet_member 
         Caption         =   "«⁄÷«¡ «·Ã„⁄Ì… «·⁄„Ê„Ì…"
      End
   End
   Begin VB.Menu mn_report 
      Caption         =   "«· ﬁ«—Ì—"
      Begin VB.Menu tm_grdpaid2 
         Caption         =   "«·«⁄÷«¡ «·⁄«„·Ì‰ Õ”» „Õ· «·«ﬁ«„…"
      End
      Begin VB.Menu tm_address2 
         Caption         =   "«·«⁄÷«¡ «·⁄«„·‰Ì Õ”» „Õ· «·«ﬁ«„…"
      End
      Begin VB.Menu tm_gridpaid1 
         Caption         =   "≈Ã„«·Ì «Ì’«·«  ”œ«œ «·√⁄÷«¡"
      End
      Begin VB.Menu tm_grdmember1 
         Caption         =   "ÿ»«⁄… »Ì«‰«  «·« ’«· »«·«⁄÷«¡"
      End
      Begin VB.Menu tm_grdmember2 
         Caption         =   "«⁄÷«¡ ”œœÊ« »œÊ‰ ﬂ«—‰ÌÂ« "
      End
      Begin VB.Menu tm_grd_phones_inv 
         Caption         =   " ·Ì›Ê‰«  «⁄÷«¡ «·œ⁄Ê…"
      End
      Begin VB.Menu tm_report 
         Caption         =   " ﬁ«—Ì— «·«⁄÷«¡"
      End
      Begin VB.Menu tm_report_ins 
         Caption         =   " ﬁ—Ì— «·⁄÷ÊÌ… «·„ﬁ”ÿ…"
      End
      Begin VB.Menu tm_grdHonor 
         Caption         =   " ﬁ—Ì— «⁄÷«¡ «·‘—›ÌÌ‰"
      End
   End
   Begin VB.Menu mn_fawary 
      Caption         =   "›Ê—Ì"
      Begin VB.Menu tm_create_fawry 
         Caption         =   "⁄„· „ÿ«·»«  ›Ê—Ì ··«⁄÷«¡ «·⁄«„·Ì‰"
      End
      Begin VB.Menu tm_fawary 
         Caption         =   "”Õ» »Ì«‰«  ›Ê—Ì ··«⁄÷«¡ «·⁄«„·Ì‰"
      End
      Begin VB.Menu tm_fawary_grid 
         Caption         =   " ﬁ—Ì— ”Õ» «·›Ê—Ì"
      End
      Begin VB.Menu sp_fawry1 
         Caption         =   "-"
      End
      Begin VB.Menu tm_create_fawry_install 
         Caption         =   "⁄„· „ÿ«·»«  ›Ê—Ì ··«⁄÷«¡ «·„ﬁ”ÿÌ‰"
      End
   End
   Begin VB.Menu mn_Services 
      Caption         =   "Œœ„« "
      Begin VB.Menu tmsecurity 
         Caption         =   "’·«ÕÌ« "
      End
      Begin VB.Menu tm_password_change 
         Caption         =   " €ÌÌ— ﬂ·„… «·”—"
      End
      Begin VB.Menu tm_back_up 
         Caption         =   "⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
      End
      Begin VB.Menu tm_send_door 
         Caption         =   "«—”«· «·»Ì«‰«  «·Ì «·»Ê«»…"
      End
      Begin VB.Menu tm_gates 
         Caption         =   "«·»Ê«»…"
      End
      Begin VB.Menu tm_printer 
         Caption         =   "÷»ÿ «·ÿ«»⁄« "
      End
      Begin VB.Menu tm_fix_member_paid 
         Caption         =   "÷»ÿ ”œ«œ «·«⁄÷«¡"
      End
      Begin VB.Menu tm_tax_install 
         Caption         =   "÷»ÿ ﬁÌ„… „÷«›… ··„ﬁ”ÿÌ‰"
      End
      Begin VB.Menu tm_tax_members_late 
         Caption         =   "÷»ÿ „ √Œ—«  «·÷—Ì»…"
      End
      Begin VB.Menu tm_year_interest 
         Caption         =   "«÷«›… „⁄œ· «·›«∆œ… ·· ﬁ”Ìÿ"
      End
      Begin VB.Menu tm_cardqry_send 
         Caption         =   "«—”«· ÿ»«⁄…"
      End
      Begin VB.Menu line_exit 
         Caption         =   "-"
      End
      Begin VB.Menu tm_Exit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub MDIForm_Load()
SetKbLayout Lang_AR
openCon con

If Not bSupermode Then LoadMenu

If Trim(LCase(RetSetting("sport", App.Path & "\conf.txt"))) = "yes" Then
    hideMenu
End If

aSeason = GetFields("Select Top 1 * from years_codes where Date1 <= " & DateSq(Date) & " and date2 >= " & DateSq(Date) & " order by date1 desc", con)
If IsEmpty(aSeason) Then
    aSeason = GetFields("Select Top 1 * from years_codes order by date1 desc", con)
End If
sSeason = retFlag(aSeason, "code")

sDate_Season = myFormat(retFlag(aSeason, "date1"))

cMsgExit = ArbString("»«·Œ—ÊÃ ” ›ﬁœ «· ⁄œÌ·«  ⁄·Ì «·”Ã· ? Â·  Êœ«·Õ›Ÿ ø")
If retFlag(aSec, "CASH") Then
    closedayfrm.Show
    Exit Sub
End If


'hideMenu
On Error GoTo myerror
'FixAddress GetCon
'FixData GetCon, "test"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
closeCon con
Set Main = Nothing
End
End Sub
Private Sub mnst_list_Click()
listFrm.Show
End Sub

Private Sub tm_accounts_Click()
acountsfrm.Show 1
End Sub

Private Sub tm_address_Click()
addressfrm.Show 1
End Sub

Private Sub tm_advance_Click()
advancefrm.Show
End Sub
Private Sub tm_arrive_day_Click()
arrivefrm2.Show 1
End Sub

Private Sub tm_arrive_month_Click()
arrivefrm1.Show 1
End Sub

Private Sub tm_address2_Click()
Addresses_ifrm.Show
End Sub

Private Sub tm_back_up_Click()
copyflashfrm.Show 1
End Sub
Private Sub tm_transdata_Click()
transDatafrm.Show 1
End Sub
Private Sub tm_cars_Click()
CarsFrm.Show
End Sub
Private Sub tm_cash_client1_Click()
cash_client1.Show
End Sub
Private Sub tm_cash_client2_Click()
cash_client2.Show
End Sub
Private Sub tm_cash_sup1_Click()
cash_sup.Show
End Sub
Private Sub tm_cur_Click()
curfrm.Show 1
End Sub

Private Sub tm_charge_service_Click()
purchase_chargefrm.Show 1
End Sub

Private Sub tm_credit_Click()
creditfrm.Show
End Sub

Private Sub tm_credit_codes_Click()
credit_codesfrm.Show 1
End Sub

Private Sub tm_credit_move_Click()
creditMovefrm.Show
End Sub

Private Sub tm_debit_Click()
debitfrm.Show
End Sub

Private Sub tm_debit_move_Click()
DebitMovefrm.Show
End Sub
Private Sub tm_driver_Click()
driverfrm.bDriver = True
driverfrm.Show 1
End Sub
Private Sub tm_emp_cash_Click()
emp_wagefrm.Show
End Sub
Private Sub tm_emp_move_Click()
empmovefrm.Show
End Sub
Private Sub tm_fair_Click()
Fairfrm.Show
End Sub

Private Sub tm_fix_codes_Click()

fix_codesfrm.Show 1
End Sub

Private Sub tm_grd_bank1_Click()
grdbankfrm1.Show
End Sub
Private Sub tm_grd_bank2_Click()
grdbankfrm2.Show
End Sub

Private Sub tm_grd_cdebit1_Click()
grdCredit1.Show
End Sub

Private Sub tm_grd_debit1_Click()
grdDebit1.Show
End Sub

Private Sub tm_gas_codes_Click()
gas_codesfrm.Show 1
End Sub

Private Sub tm_gas_orders_Click()
gas_ordersfrm.Show 1
End Sub

Private Sub tm_grdDriver1_Click()
grdDriverfrm1.Show
End Sub

Private Sub tm_grdDriver2_Click()
grdDriverfrm2.Show
End Sub

Private Sub tm_grdgas3_Click()
grdgasfrm3.Show
End Sub

Private Sub tm_grdgas5_Click()
grdgasfrm5.Show
End Sub

Private Sub tm_grditem3_Click()
grditem3.Show
End Sub

Private Sub tm_grdtravel1_Click()
grdgasfrm2.Show
End Sub
Private Sub tm_grdtravel2_Click()
grdTravelfrm2.Show
End Sub
Private Sub tm_grdtrust1_Click()
grdTrust2.Show
End Sub
Private Sub tm_loan_Click()
debit_codesfrm.Show 1
End Sub
Private Sub tm_printed_Click()
printersfrm.Show 1
End Sub

Private Sub tm_services_Click()
itemsgrdfrm.bedit = True
'itemsgrdfrm.sType = "2"
itemsgrdfrm.Show
End Sub

Private Sub tm_trans_box_move_Click()
grdtrans_box.Show
End Sub

Private Sub tm_badge_Click()
badgefrm.Show
End Sub

Private Sub tm_bankers_Click()
bankersfrm.Show
End Sub

Private Sub tm_card_honor_Click()
cardqryHonorfrm.Show
End Sub

Private Sub tm_card_service_Click()
cardqryservice.Show
End Sub

Private Sub tm_card_student_Click()
cardqryStudentfrm.Show
End Sub

Private Sub tm_card_inv_Click()
cardqry_ifrm.Show
End Sub

Private Sub tm_card_worker_Click()
cardqryworkerfrm.Show
End Sub

Private Sub tm_card1_Click()
cardqryfrm.Show
'cardqry_sendfrm.Show
End Sub

Private Sub tm_cardqry_chair_Click()
cardqrychairfrm.Show
End Sub

Private Sub tm_cardqry_send_Click()
cardqry_sendfrm.Show
End Sub

Private Sub tm_close_day_Click()
closedayfrm.Show
End Sub

Private Sub tm_door_Click()
maindoorfrm.Show 1
End Sub

Private Sub tm_create_fawry_Click()
CreatePaidFawry.Show 1
End Sub

Private Sub tm_create_fawry_install_Click()
CreatePaidFawry_install.Show 1
End Sub

Private Sub tm_exit_Click()
End
End Sub

Private Sub tm_exit_sport_Click()
End
End Sub

Private Sub tm_fawary_Click()
FawryGetPamentfrm.Show 1
End Sub

Private Sub tm_fawary_grid_Click()
grdpaid4.Show
End Sub

Private Sub tm_fix_member_paid_Click()
FixMemberPaidfrm.Show 1
End Sub

Private Sub tm_gates_Click()
maindoorfrm.Show
End Sub

Private Sub tm_grd_phones_inv_Click()
grdPhonesfrm.Show 1
End Sub

Private Sub tm_grdHonor_Click()
grdHonorfrm.Show
End Sub
Private Sub tm_grdmember1_Click()
grdmember1.Show
End Sub
Private Sub tm_grdpaid_st3_Click()
grdpaid_st3.Show
End Sub
Private Sub tm_grdpaid3_Click()
grdpaid3.Show
End Sub
Private Sub tm_gridpaid_bg1_Click()
grdpaid_bg1.Show
End Sub
Private Sub tm_gridpaid_bg2_Click()
grdpaid_bg2.Show
End Sub
Private Sub tm_gridpaid_sport1_Click()
grdpaid_sport1.Show
End Sub
Private Sub tm_gridpaid_sport2_Click()
grdpaid_sport2.Show
End Sub
Private Sub tm_gridpaid_sport3_Click()
grdpaid_sport3.Show
End Sub
Private Sub tm_gridpaid_sv1_Click()
grdpaid_sv1.Show
End Sub

Private Sub tm_gridpaid_sv2_Click()
grdpaid_sv2.Show
End Sub

Private Sub tm_grdmember2_Click()
grdmember2.Show
End Sub

Private Sub tm_grdpaid2_Click()
Addressesfrm.Show
End Sub

Private Sub tm_gridpaid1_Click()
grdpaid1.Show
End Sub
Private Sub tm_grdpaid_st1_Click()
grdpaid_st1.Show
End Sub
Private Sub tm_honor_Click()
honorfrm.Show
End Sub
Private Sub tm_item_badge_Click()
itemsfrm4.Show
End Sub
Private Sub tm_item_eng_Click()
itemsfrm3.Show
End Sub
Private Sub tm_item_sport_Click()
itemsfrm2.Show
End Sub
Private Sub tm_items_Click()
itemsfrm.bedit = True
itemsfrm.Show
End Sub
Private Sub tm_items2_Click()
paidfrm2.Show
End Sub
Private Sub tm_items3_Click()
paidfrm4.Show
End Sub
Private Sub tm_meet_member_Click()
meetMemberfrm.Show
End Sub
Private Sub tm_meeting_Click()
meetingFrm.Show
End Sub
Private Sub tm_member_Click()
memberfrm.Show
End Sub
Private Sub tm_oil_orders_Click()
oil_ordersfrm.Show 1
End Sub
Private Sub tm_travel_Click()
Ordersfrm.Show 1
End Sub
Private Sub tm_travel_trust_cash_Click()
trust_cashfrm.Show 1
End Sub
Private Sub tm_travel_trust_Click()
trustfrm.Show
End Sub
Private Sub tm_visa_Click()
grdVisafrm1.Show
End Sub
Private Sub tm_visa2_Click()
grdVisafrm2.Show
End Sub
Private Sub tm_trust_move_Click()
trustMovefrm.Show
End Sub
Private Sub tm_grditem1_Click()
grditem1.Show
End Sub
Private Sub tm_workers_Click()
driverfrm.bDriver = False
driverfrm.Show 1
End Sub
Private Sub tmBankData_Click()
bankfrm.Show
End Sub
Private Sub tmbankDtl1_Click()
rpBank3.Show
End Sub
Private Sub tmBankGrroup_Click()
ReDim aPublic(5)
aPublic(0) = "FILE5_50"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
aPublic(4) = "≈”„ «·„Ã„Ê⁄…"
aPublic(5) = "„Ã„Ê⁄«  «·»‰Êﬂ"
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmBankInout_Click()
'    bEdit = RetSec(tmItem.Name)
'    bDel = RetSec(tmItem.Name, "Del")
    'bank_in_outfrm.bEdit = True
    bank_in_outfrm.Show 1
End Sub
Private Sub tmBankItems_Click()
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sTable = "FILE5_00"
oFlagfrm.sCaption = "«ﬂÊ«œ Õ—ﬂ… «·»‰ﬂ"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.Show 1
End Sub
Private Sub tmBankMove_Click()
BankMovefrm.Show
End Sub
Private Sub tmBankMoveTotal_Click()
    rpBank2.Show
End Sub
Private Sub tmBankState_Click()
    bankStatefrm.Show
End Sub
Private Sub tmbarcodeprint_Click()
    Dream_Bar.Show
End Sub
Private Sub tmbarcode_Click()
barcodefrm.Show 1
End Sub
Private Sub tmbox_Click()
    Boxfrm.Show
End Sub
Private Sub tmboxbal_Click()
boxMovefrm.Show 1
End Sub
Private Sub tmboxtrans_Click()
    bedit = True
    boxtransfrm.Show
End Sub
Private Sub tmcharge_Click()
chargefrm1.Show 1
End Sub
Private Sub tmchargecode_Click()
'Dim oFlagGroup As New FlagGroupFrm
'oFlagGroup.sCaption = "«·„’«—Ì›"
'oFlagGroup.SCODE = "«·ﬂÊœ"
'oFlagGroup.sDesca = "«·„’—Ê›"
'oFlagGroup.sGroupDesca = "«·„’«—Ì› «·—∆Ì”Ì…"
'oFlagGroup.sTable = "FILE8_51"
'oFlagGroup.sTableGroup = "FILE8_52"
'oFlagGroup.nZero = 3
'oFlagGroup.nZeroGroup = 3
'oFlagGroup.sGroupCaption = "«·„’«—Ì› «·—∆Ì”Ì…"
'oFlagGroup.Show 1
charge_codefrm.Show 1
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
FlagFrm2.bedit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub
Private Sub tmChargerep_Click()
    rpCharge.Show
End Sub
Private Sub tmChqIn_Click()
'bEdit = RetSec(tmCollectPaper.Name)
    publicFlag = 1
    bedit = True
    chqClientfrm.Show 1
End Sub
Private Sub tmChqOut_Click()
'bEdit = True
'chqsupfrm.Show 1
checksupfrm.Show 1
End Sub
Private Sub tmChqRep_Click()
rpChq.Show
End Sub
Private Sub tmClientReport_Click()
rpClient.Show
End Sub
Private Sub tmcolor_Click()
colorfrm.bedit = True
colorfrm.Show
End Sub
Private Sub TMCustImp_Click()
CustSalesImp.Show
End Sub
Private Sub tmCut_Click()
Cutfrm.bedit = True
Cutfrm.Show
End Sub
Private Sub tmDamage_Click()
damagefrm.myPublic = 1
damagefrm.Show
End Sub
Private Sub tm_grdgas4_Click()
grdGasfrm4.Show
End Sub
Private Sub tmimpcost_Click()
impcostfrm.bedit = True
impcostfrm.Show
End Sub
Private Sub tmincome_Click()
incomefrm.Show 1
End Sub
Private Sub tmincomecode_Click()
Dim oFlagGroup As New FlagGroupFrm
oFlagGroup.sCaption = "«·«Ì—«œ« "
oFlagGroup.sCode = "«·ﬂÊœ"
oFlagGroup.sDesca = "«·«Ì—«œ« "
oFlagGroup.sGroupDesca = "«·«Ì—«œ«  «·—∆Ì”Ì…"
oFlagGroup.sTable = "FILE8_61"
oFlagGroup.sTableGroup = "FILE8_62"
oFlagGroup.nZero = 3
oFlagGroup.nZeroGroup = 3
oFlagGroup.sGroupCaption = "«·«Ì—«œ«  «·—∆Ì”Ì…"
oFlagGroup.Show 1
End Sub
Private Sub tmincomemaincode_Click()
ReDim aPublic(6)
aPublic(0) = "FILE8_61"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·«Ì—«œ"
aPublic(4) = "»Ì«‰ «·«Ì—«œ"
aPublic(5) = "«Ì—«œ«  —∆Ì”Ì…"
aPublic(6) = 3
FlagFrm2.bedit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub
Private Sub tmindust_Click()
industFrm.Show 1
End Sub
Private Sub tminput_Click()
damagefrm.myPublic = 2
damagefrm.Show
End Sub
Private Sub tmitemgroup_Click()
'itemsGroupFrm.bedit = True
'itemsGroupFrm.Show

Dim oFlagGroup As New FlagGroupFrm
oFlagGroup.sCaption = "„Ã„Ê⁄«  «·«’‰«›"
oFlagGroup.sCode = "«·ﬂÊœ"
oFlagGroup.sDesca = "≈”„ «·„Ã„Ê⁄…"
oFlagGroup.sGroupDesca = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
oFlagGroup.sTable = "FILE1_50"
oFlagGroup.sTableGroup = "FILE1_50G"
oFlagGroup.nZero = -1
oFlagGroup.nZeroGroup = -1
oFlagGroup.sGroupCaption = "„Ã„Ê⁄«  «·«’‰«› «·—∆Ì”Ì…"
oFlagGroup.Show 1

End Sub
Private Sub tmitemgroupmain_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show
End Sub

Private Sub tmItemMainGroupMain_Click()

End Sub

Private Sub tmItemMainGroupRaw_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show
End Sub

Private Sub tmoutput_Click()
outputfrm.Show
End Sub

Private Sub tmpart_Click()
partfrm.Show
End Sub
Private Sub tmPath_Click()
SettingFrm.Show 1
End Sub

Private Sub tmproduct_Click()
productfrm.bedit = True
productfrm.Show
End Sub

Private Sub tmRawitem_Click()
itemrawfrm.bedit = True
itemrawfrm.Show 1
End Sub

Private Sub tmsection_Click()
sectionfrm.bedit = True
sectionfrm.Show
End Sub
Private Sub tmpart_code_Click()
Dim oFlagGroup  As New FlagGroupFrm
oFlagGroup.sCaption = "»Ì«‰«  «·‘—ﬂ«¡"
oFlagGroup.sCode = "«·ﬂÊœ"
oFlagGroup.sDesca = "«·«”„"
oFlagGroup.sGroupDesca = "«·„Ã„Ê⁄…"
oFlagGroup.sTable = "FILE8_71"
oFlagGroup.sTableGroup = "FILE8_71G"
oFlagGroup.nZero = 3
oFlagGroup.nZeroGroup = -1
oFlagGroup.sGroupCaption = "„Ã„Ê⁄… «·‘—ﬂ«¡"
oFlagGroup.Show 1
End Sub

Private Sub tm_member_i_Click()
member_ifrm.Show
End Sub
Private Sub tm_member_inv_Click()
Member_hfrm.Show
End Sub

Private Sub tm_members_Click()
'maindoorfrm.Show 1
sportfrm.Show 1
End Sub

Private Sub tm_paid_Click()
paidfrm.Show
End Sub
Private Sub tm_paid_service_Click()
paidfrm5.Show
End Sub
Private Sub tm_paid_student_Click()
paidfrm3.Show
End Sub
Private Sub tm_paid_total_Click()
paid_totalfrm.Show
End Sub

Private Sub tm_paid_install_Click()
paid_installfrm.Show
End Sub

Private Sub tm_password_change_Click()
userfrm.Show 1
End Sub

Private Sub tm_print_card_h_Click()
cardqry_hfrm.Show
End Sub

Private Sub tm_printer_Click()
printersfrm.Show 1
End Sub

Private Sub tm_realtion_codes_Click()
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sTable = "RELATION_CODES"
oFlagfrm.sCaption = "«ﬂÊ«œ «·ﬁ—«»…"
oFlagfrm.nZero = -1
oFlagfrm.bedit = True
oFlagfrm.sWhere = "CODE <> 0"
oFlagfrm.Show 1
End Sub
Private Sub tm_service_Click()
servicefrm.Show
End Sub
Private Sub tm_service_item_Click()
itemsfrm5.Show 1
End Sub
Private Sub tm_student_Click()
studentfrm.Show
End Sub
Private Sub tm_student_old_Click()
student_oldfrm.Show
End Sub

Private Sub tm_report_Click()
reportfrm.Show
End Sub

Private Sub tm_report_ins_Click()
report_insfrm.Show
End Sub

Private Sub tm_send_door_Click()
SendDataFrm.Show
End Sub

Private Sub tm_sport_Click()
'sportfrm.Show 1
End Sub

Private Sub tm_tax_install_Click()
memberPayfrm.bedit = True
memberPayfrm.Show
End Sub

Private Sub tm_tax_members_late_Click()
FixMembeTaxfrm.Show 1
End Sub
Private Sub tm_worker_Click()
workersfrm.Show
End Sub

Private Sub tm_year_interest_Click()
Years_Interestfrm.Show 1
End Sub

Private Sub tmsecurity_Click()
Security.Show 1
End Sub
Private Sub tmshopproft_Click()
Tproft_shop.Show
End Sub
Private Sub tmStock_Click()
bedit = True
StockFrm.Show
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
About.Show 1
End Sub
Private Sub xCASH2_Click()
bedit = True
Cashfrm.Show 1
End Sub
Private Sub xExit_Click()
    End
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
Purchasefrm.bedit = True
Purchasefrm.myPublic = 0
Purchasefrm.Show
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
'salesfrm.myPublic = 0
'salesfrm.bedit = True
salesfrm.Show
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
    VsTCustSales.Show
End Sub
Private Sub xVendorData_Click()
supfrm.Show 1
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
FlagFrm2.bedit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub LoadMenu()
On Error Resume Next
For I = 0 To Main.Count - 1
    If TypeOf Main(I) Is Menu And Mid(Main(I).Name, 1, 2) = "mn" Then
        Main(I).Visible = False
    End If
Next
Err.Clear
Dim loctable As New ADODB.Recordset
loctable.Open "select Menusetting.*, menu.mainMenu from Menusetting inner join menu on menusetting.control = menu.control where menusetting.code = " & nUsercode, con, adOpenStatic, adLockReadOnly, adCmdText
For I = 0 To Main.Count - 1
    If TypeOf Main(I) Is Menu And Mid(Main(I).Name, 1, 2) <> "mn" Then
        loctable.Find "control = " & MyParn(Main(I).Name), , adSearchForward, adBookmarkFirst
        If Not loctable.EOF Then
            Main(loctable!MainMenu).Visible = True
            Main(I).Tag = IIf(loctable!Editable, "ok", "")
            Main(I).Visible = True
        Else
            Main(I).Visible = False
        End If
    End If
Next
Err.Clear
End Sub
Private Function RetMainMenu(cCode, cControl, pCon) As Variant
Dim obj As New ADODB.Recordset, aRet(1)
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("code", adInteger, adParamInput, 15, cCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("control", adVarWChar, adParamInput, 50, cControl)
cmdTable.Parameters.Append cmdTable.CreateParameter("bVisible", adInteger, adParamOutput)
cmdTable.Parameters.Append cmdTable.CreateParameter("sMainMenu", adVarWChar, adParamOutput, 25)


cmdTable.CommandText = "retMainMenu"
Set obj = cmdTable.Execute
aRet(0) = cmdTable.Parameters(2).Value
aRet(1) = cmdTable.Parameters(3).Value
Set obj = Nothing
RetMainMenu = aRet
End Function
Private Function hideMenu()
On Error Resume Next
mn_items.Visible = False
mn_meeting.Visible = False
mn_Services.Visible = False
mn_report.Visible = False
End Function
