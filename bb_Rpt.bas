Option Compare Database
Option Explicit
Dim Dbg As Boolean
Const LnkColVblzCurRate$ = _
" Sku       Txt Material      |" & _
" CurRateAc Dbl [     Amount] |" & _
" VdtFm     Txt [Valid From]  |" & _
" VdtTo     Txt [Valid to] |" & _
" HKD       Txt Unit|" & _
" Per       Txt per |" & _
" CA_Uom    Txt Uom"

Const LnkColVblzUom$ = _
 "Sku     Txt Material |" & _
 "Des     Txt [Material Description] |" & _
 "AC_U    Txt [Unit per case] |" & _
 "SkuUom  Txt [Base Unit of Measure] |" & _
 "BusArea Txt [Business Area]"
Const LnkColVblzGLBal$ = "BusArea Txt [Business Area Code] | GLBal Dbl"
Public Const LnkColVblzMB52$ = _
    " Sku    Txt Material |" & _
    " Whs    Txt Plant    |" & _
    " Loc    Txt [Storage Location] |" & _
    " BchNo  Txt Batch |" & _
    " QInsp  Dbl [In Quality Insp#]|" & _
    " QUnRes Dbl UnRestricted|" & _
    " QBlk   Dbl Blocked"

Sub Import()
'Create 5-Imp-Table [#I*] from 5-lnk-table [>*]
WImp ">GLBal", LnkColVblzGLBal
WImp ">UOM", LnkColVblzUom, "Plant='8601'"
WImp ">CurRate", LnkColVblzCurRate
WImp ">MB52", LnkColVblzMB52, "Plant='8601' and [Storage Location]='0002'"
WRun "Alter Table [#IMB52] add column OH Double"
WRun "Update [#IMB52] set OH = QUnRes+QBlk+QInsp"
WImpTbl "Permit PermitD SkuRepackMulti SkuTaxBy3rdParty SkuNoLongerTax"
End Sub
Sub MsgRunQry(A$)
MsgSet "Running query (" & A & ") ..."
End Sub
Sub Tmp()
MsgRunQry "$CurRate":    TmpCurRate
MsgRunQry "$StkOH":      TmpStkOH
MsgRunQry "$BchRate":    TmpBchRate
MsgRunQry "$RepackRate": TmpRepackRate
End Sub
Function Lnk() As String()
Dim A$(), B$(), C$(), D$(), E$(), F$(), O$()
A = WtLnkFx(">UOM", IFxUOM)
B = WtLnkFx(">MB52", IFxMB52)
C = WtLnkFx(">CurRate", IFxZHT0)
D = WtLnkFx(">GLBal", IFxGLBal)
E = WtLnkFb("Permit PermitD", IFbDuty)
F = WtLnkFb("SkuRepackMulti SkuNoLongerTax SkuTaxBy3rdParty", IFbStkHld)
O = AyAddAp(A, B, C, D, E, F)
If Sz(O) > 0 Then Lnk = O: Exit Function
A = WtChkCol(">UOM", LnkColVblzUom)
B = WtChkCol(">MB52", LnkColVblzMB52)
C = WtChkCol(">CurRate", LnkColVblzCurRate)
D = WtChkCol(">GLBal", LnkColVblzGLBal)
Lnk = AyAddAp(A, B, C, D)
End Function
Function IFbDuty$()
If IsDev Then
    IFbDuty = CurrentDb.Name
Else
    IFbDuty = "N:\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
End If
End Function

Function IFbStkHld$()
If IsDev Then
    IFbStkHld = CurrentDb.Name
Else
    IFbStkHld = "N:\SAPAccessReports\StockHolding6\StockHolding6_Data.mdb"
End If
End Function

Function Rpt()
If Not IsDev Then
    On Error GoTo X
End If
WIni
MsgClr
MsgSet "Running report...."
WOpn
If AyBrwEr(Lnk) Then Exit Function
If AyBrwEr(Er) Then Exit Function
Import
Tmp
Oup
Gen
WQuit
X:
MsgBox "Error" & vbCrLf & vbCrLf & Err.Description, vbCritical
End Function
Function Er() As String()
Er = ErzMB52_8601_0002_Missing
End Function
Private Sub ZZ_ErzMB52()
AyDmp ErzMB52
End Sub
Function ErzMB52() As String()
ErzMB52 = ErzMB52_8601_0002_Missing
End Function
Function ErzMB52_8601_0002_Missing() As String()
Dim N&, O$()
DbtLnkFx W, "#A", IFxMB52
N = DbtNRow(W, "#A", "[Plant]='8601' and [Storage Location]='0002'")
WDrp "#A"
If N = 0 Then
    Push O, "MB52 Excel: " & IFxMB52
    Push O, "Worksheet : Sheet1"
    Push O, "Above MB52 file has no data for [Plant]=8601 and [Storage Location]=0002"
    Push O, "------------------------------------------------------------------------"
    'ErzMB52_8601_0002_Missing = O
End If
End Function
Sub OBchRateH()
'Create [@BchRateHstV] [@BchRateHstV] [@BchRateFstLas]
'From #IPermit #IPermitD #OH #IUom
WDrp "#OH #H #V #FstLas #Las #Las1 #Fst #Fst1 #Sku #FstLas1 @BchRateHstV @BchRateHstH @BchRateFstLas"
'#OH
WRun "Select Distinct Sku,Sum(x.OH) as OH into [#OH] from [$StkOH] x group by Sku"

'#V & @BchRateHstV
WRun "Select a.PermitNo,a.PermitDate,x.* into [#V] from [#IPermitD] x left join [#IPermit] a on x.Permit=a.Permit order by Sku,PermitDate"
WRun "Alter Table [#V] add column BchRateSeq Int,Rate1 Long, Year Int, OH Long"
WRun "UPdate [#V] x inner join [#OH] a on x.Sku=a.Sku set x.OH = a.OH"
WRun "Update [#V] set BchRateSeq=0, Rate1 = Round(Rate,0), Year=Year(PermitDate)"
DbtUpdSeq W, "#V", "BchRateSeq", "Sku Year", "Sku Rate1"
WRun "Alter Table [#V] add column Des Text(255),Ac_U Int"
WRun "Update [#V] x inner join [#IUom] a on x.Sku=a.Sku set x.Ac_U=a.Ac_U,x.Des=a.Des"
WRun "Select * into [@BchRateHstV] from [#V]"

'#H & @BchRateHstH
WRun "Select Distinct Sku,BchRateSeq,Year(PermitDate) AS Year,Rate1,Sum(x.Qty) As Qty,Sum(x.Amt) as Amt, CLng(0) as OH" & _
" Into [#H] from [#V] x group by Sku,Year(PermitDate),BchRateSeq,Rate1"
WRun "UPdate [#H] x inner join [#OH] a on x.SKu=a.Sku set x.OH = a.OH"
WRun "Alter Table [#H] add column Des Text(255),Ac_U Int"
WRun "Update [#H] x inner join [#IUom] a on x.Sku=a.Sku set x.Ac_U=a.Ac_U,x.Des=a.Des"
WRun "Select * into [@BchRateHstH] from [#H]"

'#FstLas #Las #Las1
WRun "Select Distinct Sku,Max(x.PermitDate) as LasPermitDate,Min(x.PermitDate) as FstPermitDate into [#FstLas] from [#V] x group by Sku"
WRun "Select Distinct x.Sku,x.LasPermitDate,Max(a.Rate) as LasRate into [#Las] from [#FstLas] x inner join [#V] a on x.Sku=a.Sku and x.LasPermitDate=a.PermitDate group by x.Sku,x.LasPermitDate"
WRun "Select Distinct x.Sku,Max(a.Permit) as LasPermit,x.LasPermitDate,x.LasRate into [#Las1] from [#Las] x inner join [#V] a on x.Sku=a.Sku and x.LasPermitDate=a.PermitDate and x.LasRate=a.Rate group by x.Sku,x.LasPermitDate,x.LasRate"
'
'#Fst = Sku,FstPermitDate | FstRate
'#Fst1 = Sku,FstPermitDate,FstPermit | FstRate
WRun "Select Distinct x.Sku,x.FstPermitDate,Max(a.Rate) as FstRate into [#Fst] from [#FstLas] x inner join [#V] a on x.Sku=a.Sku and x.FstPermitDate=a.PermitDate group by x.Sku,x.FstPermitDate"
WRun "Select Distinct x.Sku,Min(a.Permit) as FstPermit,x.FstPermitDate,x.FstRate into [#Fst1] from [#Fst] x inner join [#V] a on x.Sku=a.Sku and x.FstPermitDate=a.PermitDate and x.FstRate=a.Rate group by x.Sku,x.FstPermitDate,x.FstRate"
'
'@BchRateFstLas #Sku #FstLas1 = Result: Sku Fst{Permit Date Rate} Las{..}
WRun "Select Distinct Sku into [#Sku] from (select Sku from [#Las] union select Sku from [#Fst]) x"
WRun "Select x.Sku,AC_U,Des,CLng(0) as OH into [#FstLas1] from [#Sku] x left join [#IUom] a on x.Sku=a.Sku"
WRun "Update [#FstLas1] x inner join [#OH] a on x.Sku=a.Sku set x.OH=a.OH"
WRun "Select x.Sku,AC_U,Des,OH,FstPermit,FstPermitDate,FstRate as FstBchRateU,LasPermit,LasPermitDate,LasRate As LasBchRateU" & _
" into [@BchRateFstLas] from ([#FstLas1] x left join [#Las1] a on x.Sku=a.Sku) left join [#Fst1] b on x.Sku=b.Sku"
WDrp "#OH #H #V #FstLas #Las #Las1 #Fst #Fst1 #Sku #FstLas1"
End Sub

Sub OZHT0A()
'Excess ZHT0 itm @ZHT0A.  Any Itm in ZHT0, but not in DutyPrepay Database, show the item
'Create [#A] { Sku Type } from SkuB
'App    [$CurRate]->Sku to #A
'Pivot
'Create @ZHT0A
'Create [#A] select Distinct SKU from PermitD  'Type=From Permit
'Append [#A] select Distinct SKU from [$CurRate]  type = From ZHT0
'Create [#B] pivot from [#A]
WDrp "#A #B #FmBch #FmZHT0 #Fm3p #FmNoTax #FmRepack @ZHT0A"
WRun "Select Distinct SKu into [#FmZHT0] From [$CurRate]"
WRun "Select Distinct Sku into [#FmBch] From PermitD"
WRun "Select Distinct SkuNew into [#FmRepack] From [#ISkuRepackMulti]"
WRun "Select Distinct Sku into [#Fm3p] from [#ISkuTaxBy3rdParty]"
WRun "Select Distinct Sku into [#FmNoTax] from [#ISkuNoLongerTax]"
WRun "Select Sku into [#A] from [#FmBch]"
WRun "Insert into [#A] select Sku from [#FmZHT0]"
WRun "Insert into [#A] (Sku) select SkuNew from [#FmRepack]"
WRun "Insert into [#A] select Sku from [#Fm3p]"
WRun "Insert into [#A] select Sku from [#FmNoTax]"
WRun "Select Distinct Sku," & _
" '' as IsZHT0," & _
" '' as IsTax," & _
" '' as IsRepack," & _
" '' as Is3p," & _
" '' as IsNoLongerTax" & _
" into [@ZHT0A] from [#A]"
WRun "Update [@ZHT0A] x inner join [#FmBch] a on x.Sku = a.Sku set IsTax='1'"
WRun "Update [@ZHT0A] x inner join [#FmZHT0] a on x.Sku = a.Sku set IsZHT0='1'"
WRun "Update [@ZHT0A] x inner join [#FmRepack] a on x.Sku = a.SkuNew set IsRepack='1'"
WRun "Update [@ZHT0A] x inner join [#Fm3p] a on x.Sku = a.Sku set Is3p='1'"
WRun "Update [@ZHT0A] x inner join [#FmNoTax] a on x.Sku = a.Sku set IsNoLongerTax='1'"

WRun "Alter Table [@ZHT0A] add column OH Long"
WRun "Select Sku,Sum(x.OH) as OH Into [#B] from [#IMB52] x where Whs='8601' and Loc='0002' group by Sku"
WRun "Update [@ZHT0A] x inner join [#B] a on x.Sku=a.Sku set x.OH=a.OH"

WRun "Alter Table [@ZHT0A] add column Alert Text(10), Addition Text(30)"
WRun "Update [@ZHT0A] set Alert='*Alert' where Nz(IsZHT0,'')<>'1' or Nz(IsTax,'')<>'1'"
WRun "Update [@ZHT0A] set Addition='*Ok-NoLongerTax'        where IsNoLongerTax='1' and IsZHT0='' and IsTax=''"
WRun "Update [@ZHT0A] set Addition='*Er-ExcessZHT0'         where IsZHT0='1' and IsTax='' and OH is null"
WRun "Update [@ZHT0A] set Addition='*Er-ExcessZHT0-With-OH' where IsZHT0='1' and IsTax='' and OH is not null"
WRun "Update [@ZHT0A] set Addition='*Er-MissZHT0'           where IsZHT0='' and IsTax='1'"
WRun "Update [@ZHT0A] set Addition='*Ok-Repack'             where IsRepack='1' and IsZHT0='1' and IsTax=''"
WRun "Update [@ZHT0A] set Addition='*Er-Repack'             where IsRepack='1' and Not (IsZHT0='1' and IsTax='')"

WDrp "#A #B #FmBch #FmZHT0 #Fm3p #FmNoTax #FmRepack"
End Sub
Sub OMain()
'Create @Oup
'$StkOH: Sku,HasCurRate,IsTax,IsRepack,IsTaxBy3rdParty,IsLoad,IsNoLongerTax,BchNo,OH
'Oup   @Oup{Sku BchNo OH
'       | Des SkuUom AC_U
'       | CurRateAC CurRateU
'       | BchRateU BchRateAC | UX | BchRateUX BchRateACX IsAlert BchRateSeq
WDrp "@Main #A #B"

WRun "SELECT * into [@Main] from [$StkOH] where IsImport"
WRun "Alter Table [@Main] drop column IsImport"

'AddCol --
WRun "Alter Table [@Main] Add Column" & _
" Des Text(255) ," & _
" SkuUom Text(3)," & _
" AC_U byte"
WRun "Update [@Main] x left join [#IUom] a on x.Sku=a.Sku set " & _
 "x.Des=Nz(a.Des,'')," & _
 "x.SkuUom=Nz(a.SkuUom,'')," & _
 "x.AC_U=Nz(a.AC_U,0)"

'Add Column BusArea
WRun "Alter Table [@Main] add column BusArea Text(255)"
WRun "Update [@Main] x inner join [#IUom] a on x.Sku = a.Sku set x.BusArea = a.BusArea"
 
'AddCol-CurRateAC
'AddCol-CurRateU
WRun "Alter table [@Main] Add Column" & _
" CurRateAC Double," & _
" CurRateU Long"
WRun "Update [@Main] x left join [$CurRate] a on x.Sku=a.Sku set " & _
 " x.CurRateAC=Nz(a.CurRateAC,0)," & _
 " x.CurRateU =Nz(a.CurRateU,0)"

'---
OMain_1_AddCol_RateTy
OMain_1a_RateTy_Bch
OMain_1b_RateTy_OHLas
OMain_1c_RateTy_Las
OMain_1d_RateTy_P3
OMain_1e_RateTy_Pac

OMain_2_AddCol_BchRateSeq
OMain_4a_AddCol_DifVal
OMain_4b_AddCol_IsBigDif
OMain_4c_AddCol_MustChg
OMain_4d_AddCol_IsAlert
OMain_5a_AddCol_CurRateVal_and_BchRateVal
OMain_5b_AddCol_BchRateAcX

OMain_6_UpdCol_ToBlank
WRun "Create Index Sku on [@Main] (Sku,BchNo)"
End Sub
Sub OMain_4a_AddCol_DifVal()
'AddCol-DifVal at SKU level.  Each with IsTax-SKU, has a var-amt
'It is eq to SKU-OH*CurRateU - Sum(Bch-OH * BchRateU)
'It can be use to filter or sorting to see which SKU has high impace
'--
WDrp "#A #B #C #D"
WRun "Select Distinct SKU,Sum(CurRateU*OH) as CurRateVal into [#A] from [@Main] group by SKU "
WRun "Select Distinct SKU,Sum(BchRateUX*OH) as BchRateVal into [#B] from [@Main] group by SKU"
WRun "Select Sku,CurRateVal,CDbl(0) as BchRateVal into [#C] from [#A]"
WRun "Insert into [#C] (Sku,BchRateVal) select * from [#B]"
WRun "Select Distinct Sku,Sum(x.BchRateVal) as BchRateVal,Sum(x.CurRateVal) as CurRateVal, CDbl(0) as DifVal" & _
" into [#D] from [#C] x group by Sku"
WRun "Update [#D] set DifVal = Round(BchRateVal-CurRateVal,0)"
'-- Add to [@Main]
WRun "Alter Table [@Main] add column DifVal Double"
WRun "Update [@Main] x inner join [#D] a on a.Sku=x.Sku set x.DifVal = a.DifVal"
WDrp "#A #B #C #D"
End Sub

Sub OMain_4c_AddCol_MustChg()
'WHen the Sku has only one BchRateUX and and it is not equal to CurRate, it is must-change
'#A Sku MaxBchRateSeq <- @Oup
'     Delete MaxBchRateSeq<>1
'#B (Sku,BchRateUX) from @Oup where BchRateSeq 1
'   #B1 Distinct Sku,Count(*) having Count(*)>
'   Stop if #B1 has record
'   Add Col CurRate
'   Update CurRate from Table-[$CurRate]
'#C from #B where CurRate<>BchRate
'Update @Oup->MustChg
WDrp "#A #B #C #B1"
WRun "Select Sku,Max(BchRateSeq) As MaxBchRateSeq into [#A] from [@Main] group by Sku"
WRun "Delete from [#A] where MaxBchRateSeq<>1"

WRun "Select Distinct Sku,BchRateUX into [#B] from [@Main] where Sku in (Select Sku from [#A])"
WRun "ALter Table [#B] add column CurRateU Long, MustChg YesNo"
WRun "Update [#B] x inner join [$CurRate] a on x.Sku=a.Sku set x.CurRateU = a.CurRateU"
WRun "Update [#B] set MustChg = Round(Nz(CurRateU,0),0)<>Round(Nz(BchRateUX,0),0)"

WRun "Select Distinct Sku,Count(*) into [#B1] from [#B] group By Sku having Count(*)>1"
'If DbqAny(W, "Select * from [#B1]") Then Stop

WRun "Select Sku into [#C] from [#B] where MustChg"
WRun "Alter Table [@Main] Add Column MustChg YesNo"
WRun "Update [@Main] X inner join [#C] a on x.sku=a.Sku set x.MustChg=true"
WDrp "#A #B #C #B1"
End Sub
Sub OMain_4b_AddCol_IsBigDif()
'AddCol-IsBigDif for each Abs(DifVal)>PnmVal("BigDifCtlVal")
WRun "Alter table [@Main] add column BigDifCtlVal long, IsBigDif yesno"
Dim A&: A = Val(PnmVal("BigDifCtlVal"))
WRun "Update [@Main] set BigDifCtlVal = " & A
WRun "update [@Main] set IsBigDif=True where Abs(DifVal)>" & A
End Sub

Function IFxZHT0$()
IFxZHT0 = PnmFfn("ZHT0")
End Function
Function FxFny(A$, Optional WsNm$ = "Sheet1") As String()
FxFny = ItrNy(FxCat(A).Tables(WsNm & "$").Columns)
End Function
Sub ZZ_IFnyGLBal()
D IFnyGLBal
End Sub
Sub ZZ_IFnyMB52()
D IFnyMB52
End Sub
Function NewFxCn() As adodb.Connection
Set NewFxCn = FxCn(TmpFx)
End Function
Function IFnyGLBal() As String()
IFnyGLBal = FxFny(IFxGLBal)
End Function

Function IFnyMB52() As String()
IFnyMB52 = FxFny(IFxMB52)
End Function

Function IFxGLBal$()
IFxGLBal = PnmFfn("GLBal")
End Function
Function OupFx$()
Dim A$, B$
A = OupPth & Apn & " " & StkDteYYYYMMDD & ".xlsm"
B = FfnNxt(A)
OupFx = B
End Function

Private Sub MsgSet(A$)
Form_Main.MsgSet A
End Sub
Private Sub MsgClr()
Form_Main.MsgClr
End Sub
Sub Gen()
MsgSet "Export to Excel ....."
OupFx_Gen OupFx, WFb, "FmtMB52B", "FmtCtl"
End Sub
Sub ZZ_MaxBchRateSeq()
Debug.Assert 2 = MaxBchRateSeq
End Sub
Sub ZZ_FmtMB52B()
Dim Wb As Workbook, Ws As Worksheet
Set Wb = TpWb
Set Ws = WbWsCd(Wb, "WsOMB52B")
WbVis Wb
FmtMB52B Wb
End Sub
Function FmlSumOHFml$(N%)
Dim Ay$(), J%, S$
If N > 5 Then Stop
For J = 1 To N
    Push Ay, "[OH" & J & "]"
Next
FmlSumOHFml = "=" & Join(Ay, "+")
End Function
Function FmlNewrateAcSel_ChooseFml$(N%)
Dim Ay$(), J%, S$
If N > 5 Then Stop
'LoSetFml Lo, "NewRateAC", "=IF(ISBLANK([@[Sel]]),,CHOOSE([@[Sel]],[@BchRateAc1],[@BchRateAc2]))"
For J = 1 To N
    Push Ay, "[@BchRateAc" & J & "]"
Next
S = Join(Ay, ",")
FmlNewrateAcSel_ChooseFml = "=IF(ISBLANK([@[Sel]]),,CHOOSE([@[Sel]]," & S & "))"
End Function
Property Get MaxBchRateSeq%()
'From @Main
MaxBchRateSeq = DbqV(W, "Select Max(BchRateSeq) from [@Main]")
End Property
Sub FmtCtl(A As Workbook)
Dim Ws As Worksheet
Dim N%, Lo As ListObject, F1$, F2$
Set Ws = WbWsCd(A, "WsCtl")
Set Lo = WsFstLo(Ws)
LoSetFml Lo, "Aft", "=SumIf(TblMB52B,[@BusArea],TblMB52B[Aft])"
LoSetFml Lo, "Bef", "=SumIf(TblMB52B,[@BusArea],TblMB52B[Bef])"
LoSetFml Lo, "Dif Aft", "=[Aft]-[GLBal]"
LoSetFml Lo, "Dif Bef", "=[Bef]-[GLBal]"
Ws.Range("$B:$F").EntireColumn.ColumnWidth = 15
End Sub

Function FmlSug(N%)
Dim O$
Dim A$, C$
Dim Ay$(), J%
Dim B$
A = "=Suggest(PlusOrMinus,[@CurRateAc],?)"
ReDim Ay(N - 1)
For J = 0 To N - 1
    Ay(J) = "[@BchRateAc" & J + 1 & "]"
Next
B = JnComma(Ay)
C = FmtQQ(A, B)
FmlSug = C
End Function

Function FmlDifVal(N%)
Dim S1$, S2$(), S3$, S4$, J%
S1 = "[OH?]*[BchRateU?]"
For J = 1 To N
    Push S2, FmtQQ(S1, J, J)
Next
S3 = Join(S2, "-")
S4 = FmtQQ("=[CurRateAc]*[OH_AC]-?", S3)
FmlDifVal = S4
End Function

Sub ZZ_FmlDifVal()
Dim A$
A = FmlDifVal(2)
Stop
End Sub
Sub FmtMB52B(A As Workbook)
Dim Ws As Worksheet
Set Ws = WbWsCd(A, "WsOMB52B")

Dim N%, Lo As ListObject, F1$, F2$, F3$, F4$
N = Min(MaxBchRateSeq, 5)
F1 = FmlSumOHFml(N)
F2 = FmlNewrateAcSel_ChooseFml(N)
F3 = FmlDifVal(N)
F4 = FmlSug(N)
Set Lo = WsFstLo(Ws)
LoSetFml Lo, "Aft", "=[@[OH_AC]]*[@[CurRateAc]]" 'In the Tp
LoSetFml Lo, "Bef", "=[@[OH_AC]]*[@[NewRateAc]]"
LoSetFml Lo, "OH", F1
LoSetFml Lo, "OH_AC", "=[@[OH]]/[@[AC_U]]"
LoSetFml Lo, "NewRateAc", "=IF([@NewRateAcSel]=0,[@CurRateAc],[@NewRateAcSel])"
LoSetFml Lo, "NewRateAcSel", F2
LoSetFml Lo, "DifVal", F3
LoSetFml Lo, "IsBigDif", "=[DifVal]>DifCtl"
LoSetFml Lo, "Suggest", F4
'---
'Delete Excess BchRate Bucket from (N+1 to 5)
Dim R As Range
If N < 5 Then
    Dim ColOffSet As Range
    Set ColOffSet = Ws.Columns("X")
    Dim R1 As Range, R2 As Range, R3 As Range
    Set R1 = RgCC(ColOffSet, N + 1, 5).EntireColumn
    Set R2 = RgCC(ColOffSet, N + 1 + 5, 5 + 5).EntireColumn
    Set R3 = RgCC(ColOffSet, N + 1 + 10, 5 + 10).EntireColumn
    'Debug.Print R1.Address
    'Debug.Print R2.Address
    'Debug.Print R3.Address
    R3.Delete
    R2.Delete
    R1.Delete
End If
Ws.Range("X1:AL3").MergeCells = False
FmtMB52B_X_1ToN_BchRateAc_Qty_U Ws.Range("X8"), N
End Sub
Function IFxMB52$()
IFxMB52 = PnmFfn("MB52")
End Function
Function IFxUOM$()
IFxUOM = PnmFfn("UOM")
End Function

Sub OMain_1_AddCol_RateTy()
WRun "Alter Table [@Main] add column RateTy text(10)"
'*Bch    - Sku+BchNo can lookup for rate
'*OHLas  - There is OH, use the last OH-Batch-Rate
'*Las    - There is no OH, use MB52-Date to lookup
'*Pac
'*3p
End Sub
Sub OMain_1a_RateTy_Bch()
WRun "Alter Table [@Main] add column" & _
" BchPermit Long, BchPermitD Long, BchPermitDate Date, BchPermitDateEnd Date," & _
" BchRateU Currency, BchRateUX Currency"
WRun "Update [@Main] x inner join [$BchRate] a on x.Sku=a.Sku and x.BchNo=a.BchNo" & _
" set" & _
" x.BchPermit=a.Permit," & _
" x.BchPermitD=a.PermitD," & _
" x.BchPermitDate=a.PermitDate," & _
" x.BchPermitDateEnd=a.PermitDateEnd," & _
" x.BchRateU=a.BchRateU," & _
" x.BchRateUX = a.BchRateU," & _
" x.RateTy='*Bch'"
End Sub
Sub OMain_1b_RateTy_OHLas()
'For those Sku with BchRateU but cannot find rate from $BchRate table, use *OHLas approach
WDrp "#A #B #Sku$ #Sku0 #Sku"
WDrp "#K2 #K3 #OHLas"
WDrp "#MB52A"
WRun "Select distinct Sku into [#Sku0] from [@Main] where Nz(BchRateU,0)=0"
WRun "Select distinct Sku into [#Sku$] from [@Main] where BchRateU>0"
WRun "Select * into [#A] from [#Sku0]"
WRun "Insert into [#A] select Sku from [#Sku$]"
WRun "Select Distinct Sku,Count(*) as Cnt into [#B] from [#A] group by Sku"
WRun "Select Sku into [#Sku] from [#B] where Cnt=2"

WRun "Select * into [#MB52A] from [@Main] where Sku in (Select * from [#Sku]) and BchRateU>0"

WRun "Select Distinct Sku,Max(x.BchPermitDate) as BchPermitDate into [#K2] from [#MB52A] x Group By Sku"
WRun "Select Distinct x.Sku,x.BchPermitDate,Max(a.BchPermitD) as BchPermitD into [#K3] from [#K2] x inner join [#MB52A] a" & _
" on x.Sku=a.Sku and x.BchPermitDate=a.BchPermitDate group By x.Sku,x.BchPermitDate"
WRun "Select a.* into [#OHLas] from [#K3] x inner join [#MB52A] a on x.Sku=a.Sku and x.BchPermitDate=a.BchPermitDate and x.BchPermitD=a.BchPermitD"

WRun "Alter Table [@Main] add column OHLas text(1)," & _
" OHLasPermitDate Date, OHLasPermitDateEnd Date," & _
" OHLasPermit Long,OHLasPermitD Long," & _
" OHLasRateU Currency, OHLasBchNo text(20)"

WRun "Update [@Main] x inner join [#OHLas] a on x.Sku=a.Sku and x.BchPermitDate=a.BchPermitDate and x.BchPermitD=a.BchPermitD" & _
" set OHLas='@' where x.Sku in (Select Sku from [#Sku])"

WRun "Update [@Main] x inner join [#OHLas] a on x.Sku=a.Sku" & _
" set" & _
" OHLas='*'," & _
" x.RateTy='*OHLas'," & _
" x.OHLasRateU = a.BchRateU," & _
" x.BchRateU = a.BchRateU," & _
" x.BchRateUX = a.BchRateUX," & _
" x.OHLasPermitDate = a.BchPermitDate," & _
" x.OHLasPermitD = a.BchPermitD," & _
" x.OHLasPermitDateEnd = a.BchPermitDateEnd," & _
" x.OHLasPermit = a.BchPermit," & _
" x.OHLasBchNo = a.BchNo" & _
" where Nz(x.BchRateU,0)=0 and x.Sku in (Select Sku from [#Sku])"
WDrp "#A #B #Sku$ #Sku0 #Sku"
WDrp "#K2 #K3 #OHLas"
WDrp "#MB52A"
End Sub

Sub OMain_1c_RateTy_Las()
'1049427
'If still no rate, try *Las-Bch: Sku + MB52-Dte to lookup $BchRate
WRun "Alter Table [@Main] add column" & _
" LasPermit Long," & _
" LasPermitD Long," & _
" LasPermitDate Date," & _
" LasPermitDateEnd Date," & _
" LasBchNo Text(20)," & _
" LasRateU Currency"

WRun FmtQQ("Update [@Main] x inner join [$BchRate] a on x.Sku=a.Sku" & _
" Set RateTy='*Las', x.BchRateUX=a.BchRateU, x.LasRateU = x.BchRateU," & _
" x.LasPermit = a.Permit," & _
" x.LasPermitD = a.PermitD," & _
" x.LasPermitDate = a.PermitDate," & _
" x.LasPermitDateEnd = a.PermitDateEnd," & _
" x.LasBchNo = a.BchNo" & _
" where x.RateTy is Null and #?# between PermitDate and PermitDateEnd", StkDteYYYYMMDD)
End Sub

Sub OMain_1e_RateTy_Pac()
'For those Sku IsRepack, use Sku to lookup from the $Repack5 (Pk=Sku)
WRun "Alter Table [@Main] add column PackRateU Currency"
WRun "Update [@Main] x inner join [@Repack5] a on x.Sku=a.SkuNew" & _
" set" & _
" RateTy='*Pac'," & _
" x.BchRateUX=a.PackRateU," & _
" x.PackRateU=a.PackRateU " & _
" where IsRepack"
End Sub
Sub OMain_1d_RateTy_P3()
'For those Sku Is3p, use Sku to lookup the rate from SkuTaxBy3rdParty
WRun "Alter Table [@Main] add column TaxBy3pRateU Currency"
WRun "Update [@Main] x inner join [#ISkuTaxBy3rdParty] a on x.Sku=a.Sku" & _
" set" & _
" RateTy='*3p'," & _
" BchRateUX=a.RateU," & _
" TaxBy3pRateU=a.RateU " & _
" where Is3p"
End Sub

Sub OMain_2_AddCol_BchRateSeq()
'Notes: BchRateSeq is running from for SKU
'Sort Sku,PermitDate (Put 2099 if is null), Round(BchRateUX)
'Set BchRateSeq
'Update back to @Main
WDrp "#A #B"

WRun "Alter Table [@Main] add column Id Int, BchRateSeq Int"
DbtUpdSeq W, "@Main", "Id"
WRun "Select Id,Sku,BchPermitDate,Round(x.BchRateUX,0) as BchRateUX into [#A] from [@Main] x"
WRun "Update [#A] set BchPermitDate=#2099/12/31# where BchPermitDate is null"

WRun "Select x.*,CInt(0) as BchRateSeq into [#B] from [#A] x order by Sku,BchPermitDate,BchRateUX"
DbtUpdSeq W, "#B", "BchRateSeq", "Sku", "BchRateUX"
WRun "Update [@Main] x inner join [#B] a on a.Id=x.Id set x.BchRateSeq = a.BchRateSeq"
WDrp "#A #B"
WRun "Alter Table [@Main] drop column Id"
End Sub
Sub OMain_5a_AddCol_CurRateVal_and_BchRateVal()
WRun "Alter Table [@Main] Add Column CurRateVal Currency,BchRateVal Currency, ValDif currency"
WRun "Update [@Main] set CurRateVal = CurRateAC * OH / AC_U, BchRateVal = BchRateUX * OH"
WRun "Update [@Main] set ValDif = CurRateVal - BchRateVal"
End Sub
Sub OMain_5b_AddCol_BchRateAcX()
WRun "Alter Table [@Main] Add Column BchRateAcX Currency"
WRun "Update [@Main] set BchRateAcX = BchRateUX * AC_U"
End Sub
Sub OMain_6_UpdCol_ToBlank()
WRun "Update [@Main] set RateTy='' where RateTy is null"
WRun "Update [@Main] set BchRateUX=0 where BchRateUX is null"
WRun "Update [@Main] set BchRateAcX=0 where BchRateAcX is null"
WRun "Update [@Main] set OHLas='' where OHLas is null"
End Sub

Sub OMain_4d_AddCol_IsAlert()
'Create [#A] from [@Main] Dist Sku Rate:CurRateU  from [@Main]
'App    [#A] from [@Main] Dist SKU Rate:BchRateUX from [@Main]
'Create [#B] from [#A]   Dist SKU Rate
'Create [#C] from [#B]   Dist SKU with Count(*)>1
'AddCol [@Main] IsAlert
'Update [@Main]->IsAlert from [#C]
WDrp "#A #B #C"
WRun "Select Distinct Sku,CurRateU as Rate into [#A] from [@Main]"
WRun "Insert into [#A] (Sku,Rate) select Distinct Sku,BchRateUX from [@Main]"
WRun "Select Distinct Sku,Rate into [#B] from [#A]"
WRun "Select Distinct Sku into [#C] from [#B] Group By Sku having Count(*)>1"
WRun "Alter Table [@Main] add column IsAlert YesNo"
WRun "Update [@Main] set IsAlert = True where Sku in (Select Sku from [#C])"
WDrp "#A #B #C"
End Sub
Function Vdt() As Boolean
Dim O$()
If Sz(O) = 0 Then Vdt = True: Exit Function
AyBrw O
End Function
Sub OMain_2_AddCol_BchRateSeq_UpdCol()
'Inp: [#A] : Sku BchNo BchRateUX BchRateSeq
'Upd: [BchRateSeq]
Const Sql$ = "Select * from [#A] order by Sku,BchNo,BchRateUX"
Dim LasSku$, LasBchRateUX&, Seq%
With W.OpenRecordset(Sql)
    While Not .EOF
        If LasSku <> !Sku And LasBchRateUX <> !BchRateUX Then
            Seq = 1
            LasSku = !Sku
            LasBchRateUX = !BchRateUX
        Else
            If LasBchRateUX <> !BchRateUX Then
                LasBchRateUX = !BchRateUX
                Seq = Seq + 1
            End If
        End If
        .Edit
        !BchRateSeq = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub
Sub Cmp_ISkuB_IPermitD()
'Compare the SKU
WDrp "#A #B #C #D"
WRun "Select Distinct SKU into [#A] from [#ISkuB]"
WRun "Select Distinct SKU into [#B] from [#IPermitD]"
WRun "Select Distinct SKU into [#C] from (Select * from [#A] union Select * from [#B])"
WRun "Select x.Sku,IIF(IsNull(a.Sku),'','X') as FndInSkuB,IIF(IsNull(b.SKU),'','X') as FndInPermitD,False as Alter into [#D]" & _
" from ([#A] x" & _
" left join [#B] a on x.Sku=a.Sku)" & _
" left join [#C] b on x.Sku=b.Sku"
WRun "Update [#D] set Alter=true where FndInSkuB='' or FndInPermitD=''"
WBrw
Stop
WDrp "#A #B #C #D"
End Sub
Sub TmpStkOH()
WDrp "$StkOH #A #B #C #D #E"
WRun "Select Distinct Sku,BchNo,Sum(x.OH) as OH into [$StkOH] from [#IMB52] x group by Sku,BchNo having Sum(x.OH)<>0"
WRun "Alter Table [$StkOH] add column" & _
" HasCurRate    YesNo," & _
" IsTax         YesNo," & _
" IsRepack      YesNo," & _
" Is3p          YesNo," & _
" IsNoLongerTax YesNo," & _
" IsImport      YesNo"
WRun "Select Distinct Sku into [#A] from [$CurRate]"
WRun "Select Distinct Sku into [#B] from [#IPermitD]"
WRun "Select Distinct SkuNew as Sku into [#C] from [#ISkuRepackMulti]"
WRun "Select Sku into [#D] from [#ISkuTaxby3rdParty]"
WRun "Select Sku into [#E] from [#ISkuNoLongerTax]"
WRun "Update [$StkOH] x inner join [#A] a on x.Sku = a.Sku set x.HasCurRate = true"
WRun "Update [$StkOH] x inner join [#B] a on x.Sku = a.Sku set x.IsTax =true"
WRun "Update [$StkOH] x inner join [#C] a on x.Sku = a.Sku set x.IsRepack= true"
WRun "Update [$StkOH] x inner join [#D] a on x.Sku = a.Sku set x.Is3p=True"
WRun "Update [$StkOH] x inner join [#E] a on x.Sku = a.Sku set x.IsNoLongerTax=True"
WRun "Update [$StkOH] set IsImport = (HasCurRate or IsTax or IsRepack or Is3p and Not IsNoLongerTax)"
WDrp "#A #B #C #D #E"
End Sub

Sub TmpBchRate()
WDrp "$BchRate #A #A1"
WRun "select Distinct Sku into [#A1] from [$StkOH]"
WRun "Insert into [#A1] (Sku) select Distinct SkuFm from [#ISkuRepackMulti]"
WRun "select Distinct Sku into [#A] from [#A1]"
WRun "select Sku,PermitDate,PermitDate as PermitDateEnd,BchNo,Rate as BchRateU,PermitD,x.Permit" & _
" into [$BchRate] from [#IPermitD] x" & _
" inner join [#IPermit] a on x.Permit=a.Permit" & _
" where Sku in (Select Distinct Sku from [#A])" & _
" Order By Sku,PermitDate"
DbtUpdToDteFld W, "$BchRate", "PermitDateEnd", "Sku", "PermitDate"
WRun "Alter Table [$BchRate] add column BchRateAc currency,Ac_U Int"
WRun "Update [$BchRate] x inner join [#IUom] a on x.Sku=a.Sku set x.AC_U=a.AC_U"
WRun "Update [$BchRate] set BchRateAC = BchRateU*AC_U where AC_U<>0"
WRun "Create Index Sku on [$BchRate] (Sku,BchNo)"
'
WDrp "#A #A1"
End Sub

Property Get StkDteYYYYMMDD$()
StkDteYYYYMMDD = Mid(PnmVal("MB52Fn"), 6, 10)
End Property

Property Get StkDte() As Date
StkDte = StkDteYYYYMMDD
End Property
Sub TmpRepackRate()
'Input:
'#Inp: Sku,Dte |
'$BchRate: PermitDate PermitDateEnd Permit PermitD BchNo Sku BchRateU
'#ISkuRepackMulti: SkuNew SkuFm FmSkuQty
Dim Dte As Date
Dte = StkDte
WDrp "#Inp"
WRun FmtQQ("Select Distinct SkuNew As SKu, #?# as Dte into [#Inp] from [#ISkuRepackMulti]", Dte)

WDrp "#A #B #C #D #E $Repack1 $Repack2 $Repack3 $Repack4 $Repack5"

WRun "Select Distinct Sku,Dte into [#A] from [#Inp]"

WRun "Select SkuNew,Dte,SkuFm,FmSkuQty into [#B] from [#A] x inner join [#ISkuRepackMulti] a on x.Sku=a.SkuNew"
WRun "Create Index Pk on [#B] (SkuNew,SkuFm)"

WRun "Select SkuNew,SkuFm,Dte,FmSkuQty,BchRateU,PermitDate,BchNo,Permit,CLng(a.PermitD) as PermitD" & _
    " into [#C] from [#B] x,[$BchRate] a where false"
WRun "Insert into [#C] select * from [#B]"
WRun FmtQQ("Update [#C] x inner join [$BchRate] a on x.SkuFm=a.Sku set " & _
" x.BchRateU  =a.BchRateU  ," & _
" x.PermitDate=a.PermitDate," & _
" x.BchNo     =a.BchNo     ," & _
" x.Permit    =a.Permit    ," & _
" x.PermitD   =a.PermitD    " & _
" where Dte between a.PermitDate and a.PermitDateEnd")
WRun "Alter Table [#C] add column Amt Currency"
WRun "Update [#C] set Amt=FmSkuQty*BchRateU"
WRun "Create Index Pk on [#C] (SkuNew,SkuFm)"

WRun "Select Distinct SkuNew,Sum(x.Amt) as PackRateU, Count(*) as FmSkuCnt,Sum(x.FmSkuQty) as FmQty into [#D] from [#C] x group by SkuNew"
WRun "Alter Table [#D] add column Ac_U integer, PacRateAc Currency"
WRun "Update [#D] x inner join [#IUom] a on x.SkuNew=a.Sku set x.Ac_U = a.Ac_U"
WRun "Update [#D] Set PacRateAc = PackRateU * Ac_U"
WRun "Create Index Pk on [#D] (SkuNew)"

WDrp "@Repack1 @Repack2 @Repack3 @Repack4 @Repack5"
WRun "Select * into [@Repack1] from [#ISkuRepackMulti]"
WRun "Select * into [@Repack2] from [#Inp]"
WtRen "#B", "@Repack3", True
WtRen "#C", "@Repack4"
WtRen "#D", "@Repack5"
WDrp "#A #Inp"
End Sub

Sub WtRen(Fmt$, ToT$, Optional ReOpnFst As Boolean)
DbtRen W, Fmt, ToT, ReOpnFst
End Sub

Sub WReOpn()
WCls
WOpn
End Sub

Sub WCls()
On Error Resume Next
W.Close
Set W = Nothing
End Sub

Sub Oup()
MsgRunQry "@CurRate":  OCurRate
MsgRunQry "@BchRate":  OBchRate
MsgRunQry "@Sku":      OSku
MsgRunQry "@Main":    OMain
MsgRunQry "@MB52B":    OMB52B
MsgRunQry "@ZHT0A":    OZHT0A
MsgRunQry "@BchRateH": OBchRateH
MsgRunQry "@GLBal":    OGLBal
End Sub
Sub OMB52B()
WDrp "@MB52B"
WDrp "#A #B #B1 #B2 #B3 #B4 #B5 @MB52B"

WRun "Select BusArea into [@MB52B] from [@Main] where False"
WRun "Alter Table [@MB52B] add column" & _
" Bef Currency," & _
" Aft Currency," & _
" OH Currency," & _
" OH_AC Currency," & _
" SkuUom Text(10)," & _
" Ac_U Integer," & _
" HasCurRate YesNo, IsTax YesNo,IsRepack YesNo,Is3p YesNo, IsNoLongerTax YesNo," & _
" IsAlert YesNo," & _
" DifVal  Long," & _
" IsBigDif YesNo," & _
" MustChg YesNo," & _
" Sku Text(255), Des Text(255)," & _
  "CurRateAc Currency, NewRateAc Currency, NewRateAcSel Currency, Sel Byte, Suggest Byte," & _
  "BchRateAc1 Currency, BchRateAc2 Currency, BchRateAc3 Currency, BchRateAc4 Currency, BchRateAc5 Currency," & _
  "OH1 Long, OH2 Long,OH3 Long,OH4 Long, OH5 Long," & _
  "BchRateU1 Currency, BchRateU2 Currency, BchRateU3 Currency, BchRateU4 Currency, BchRateU5 Currency"

WRun "Insert Into [@MB52B] (BusArea,Sku,Des,SkuUom,Ac_U,IsAlert,MustChg)" & _
" select Distinct BusArea,Sku,Des,SkuUom,Ac_U,IsAlert,MustChg" & _
" from [@Main]"
WRun "Update [@MB52B] set HasCurRate = true    where sku in (Select Sku from [@Sku] where HasCurRate)"
WRun "Update [@MB52B] set IsTax = true         where sku in (Select Sku from [@Sku] where IsTax)"
WRun "Update [@MB52B] set IsRepack = true      where sku in (Select Sku from [@Sku] where IsRepack)"
WRun "Update [@MB52B] set Is3p = true          where sku in (Select Sku from [@Sku] where Is3p)"
WRun "Update [@MB52B] set IsNoLongerTax = true where sku in (Select Sku from [@Sku] where IsNoLongerTax)"

WRun "Update [@MB52B] x inner join [@CurRate] a on x.Sku=a.Sku set x.CurRateAc=a.CurRateAc"

'--- Cpy [@Main] into [#B] & set BchRateSeq=5 for those >5
WRun "Select Sku,OH,BchRateSeq,BchRateU,BchRateUX,BchRateAcX into [#A] from [@Main] "
WRun "Update [#A] set BchRateSeq=5 where BchRateSeq>5" '-- Set BchRateSeq to 5 for those >5

'-- Create #B, #B1..#B5
WRun "Select Distinct Sku,BchRateSeq,Sum(x.OH) as OH,Max(x.BchRateUX) as BchRateU, Max(x.BchRateAcX) as BchRateAc" & _
" into [#B] from [#A] x group by Sku,BchRateSeq"

'-- AddCol BchRateU.. & OH..
Dim N% ' MaxBchRateSeq
N = DbqLng(W, "Select Max(BchRateSeq) from [#B]")
If N > 5 Then Stop ' At most =5

'-- Create #B1..#B5 from #B
Dim J%
For J = 1 To 5
    WRun FmtQQ("Select * into [#B?] from [#B] where BchRateSeq=?", J, J)
Next


'-- Update @MB52B->[OH? BchRateU? BchRateAc?] by [#B?]
Dim S$
For J = 1 To 5
    S = FmtQQ("Update [@MB52B] x inner join [#B?] a on x.Sku=a.Sku" & _
        " set x.OH?=a.OH, x.BchRateU?=a.BchRateU , x.BchRateAC?=a.BchRateAC where a.BchRateSeq=?", J, J, J, J, J)
    WRun S
Next
WDrp "#A #B #B1 #B2 #B3 #B4 #B5"
End Sub

Sub OSku()
WDrp "@Sku"
WRun "Select Distinct Sku,HasCurRate,IsTax,IsRepack,Is3p,IsNoLongerTax,IsImport into [@Sku] from [$StkOH]"
End Sub

Sub OBchRate()
WDrp "@BchRate"
WRun "Select * into [@BchRate] from [$BchRate]"
End Sub
Sub OCurRate()
'Alert for CA_Uom<>CA or Per<>1 or HKD<>HKD
WDrp "@CurRate"
WRun "Select * into [@CurRate] from [$CurRate]"
WRun "Alter Table [@CurRate] Add Column Alert Text(20)"
WRun "Update [@CurRate] set Alert='Error of HKD / CA_Uom = CA / Per = 1' where HKD<>'HKD' or CA_Uom<>'CA' or Per<>1"
End Sub
Sub OGLBal()
WDrp "@GLBal"
WRun "Select BusArea,CCur(0) as Bef,CCur(0) as Aft,GLBal,CCur(0) as [Dif Bef],CCur(0) as [Dif Aft] into [@GLBal] from [#IGLBal]"
End Sub

Sub WDrp(Tny0)
DbDrpTbl W, Tny0
End Sub
Sub QClsTbl()
AcsClsTbl WAcs
End Sub
Sub TmpCurRate()
'VdtFm & VdtTo format DD.MM.YYYY
'1: #ICurRate VdtFm VdtTo Sku CurRateAc HKD CA_Uom Per
'2: #IUom     SKu Ac_U
'O: $CurRate  Sku CurRateU AC_U CurRateU HKD CA_Uom Per
WDrp "#Cpy #Active $CurRate"
WRun "Select * into [#Cpy] from [#ICurRate]"
WRun "Alter Table [#Cpy] Add Column VdtFmDte Date,VdtToDte Date,IsCur YesNo"
WRun "Update [#Cpy] Set" & _
" VdtFmDte = DateSerial(RIGHT(VdtFm,4),MID(VdtFm,4,2),LEFT(VdtFm,2))," & _
" VdtToDte = DateSerial(RIGHT(VdtTo,4),MID(VdtTo,4,2),LEFT(VdtTo,2))"
WRun "Update [#Cpy] set IsCur = true where Now between VdtFmDte and VdtToDte"

WRun "Select Sku,CurRateAC,CByte(0) as AC_U, CLng(0) as CurRateU,HKD,CA_Uom,Per into [#Active] from [#Cpy]"
WRun "Update [#Active] x inner join [#IUom] a on x.Sku = a.Sku set x.AC_U=a.AC_U"
WRun "Update [#Active] set CurRateU = Round(CurRateAC / AC_U,4) where AC_U<>0"

WReOpn
WtRen "#Active", "$CurRate"
WDrp "#Cpy #Acitve"
End Sub

Sub DocUOM()
'InpX: [>UOM]     Material [Base Unit of Measure] [Material Description] [Unit per case]
'Oup : UOM        Sku      SkuUOM                 Des                    AC_U

'Note on [Sales text.xls]
'Col  Xls Title            FldName     Means
'F    Base Unit of Measure SkuUOM      either COL (bottle) or PCE (set)
'J    Unit per case        AC_U        how many unit per AC
'K    SC                   SC_U        how many unit per SC   ('no need)
'L    COL per case         AC_B        how many bottle per AC
'-----
'Letter meaning
'B = Bottle
'AC = act case
'SC = standard case
'U = Unit  (Bottle(COL) or Set (PCE))

' "SC              as SC_U," & _  no need
' "[COL per case]  as AC_B," & _ no need
End Sub

Sub IMB52Opn()
FxOpn IFxMB52
End Sub

Function IMB52Fny() As String()
AyDmp DbtFny(W, ">MB52")
End Function

Sub ZZ_FmtMB52B_X_1ToN_BchaRateAc_Qty_U()
FmtMB52B_X_1ToN_BchRateAc_Qty_U RgVis(NewA1), 3
End Sub

Sub FmtMB52B_X__FmtLblLin(A As Range, Txt)
FmtMB52B_X__SetColrAndBdr A
A.Merge
A.HorizontalAlignment = xlCenter
A.Value = Txt
End Sub

Function CvRg(A) As Range
Set CvRg = A
End Function

Sub FmtMB52B_X__FmtLin3(A As Range, N%)
FmtMB52B_X__SetColrAndBdr A
Dim J%
For J = 1 To N
    With CvRg(A.Cells(1, J))
        .Value = J
        .HorizontalAlignment = True
    End With
Next
End Sub
Sub FmtMB52B_X__SetColrAndBdr(A As Range)
A.Interior.Color = 65535
With A
    With .Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With .Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With .Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With .Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With .Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With .Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End With
End Sub
Sub FmtMB52B_X_1ToN_BchRateAc_Qty_U(At As Range, N%)
Dim RLin1 As Range
Dim RLin2Ac As Range
Dim RLin2Q As Range
Dim RLin2U As Range

Dim RLin3U As Range
Dim RLin3Ac As Range
Dim RLin3Q As Range

Set RLin1 = RgRCC(At, 1, 1, N)
Set RLin2Ac = RgRCC(At, 2, 1, N)
Set RLin2Q = RgRCC(At, 2, 1 + N, N + N)
Set RLin2U = RgRCC(At, 2, 1 + N + N, N + N + N)

Set RLin3Ac = RgRCC(At, 3, 1, N)
Set RLin3Q = RgRCC(At, 3, 1 + N, N + N)
Set RLin3U = RgRCC(At, 3, 1 + N + N, N + N + N)

FmtMB52B_X__FmtLblLin RLin1, "From new to old"

FmtMB52B_X__FmtLblLin RLin2Ac, "Rate (HKD/Ac)"
FmtMB52B_X__FmtLblLin RLin2Q, "Qty (Unit)"
FmtMB52B_X__FmtLblLin RLin2U, "Rate (HKD/Unit)"

FmtMB52B_X__FmtLin3 RLin3Ac, N
FmtMB52B_X__FmtLin3 RLin3Q, N
FmtMB52B_X__FmtLin3 RLin3U, N
End Sub