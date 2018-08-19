Option Compare Database
Option Explicit
Public Fso As New Scripting.FileSystemObject
Type AttRs
    TblRs As DAO.Recordset
    AttRs As DAO.Recordset
End Type
Type XlsLnkInf
    IsXlsLnk As Boolean
    Fx As String
    WsNm As String
End Type
Public W As Database
Const PSep$ = " "
Const PSep1$ = " "

Property Get Apn$()
Static X As Boolean, Y$
If Not X Then
    X = True
    Y = SqlV("Select Apn from PgmPrm")
End If
Apn = Y
End Property
Sub TmpHomBrw()
PthBrw TmpHom
End Sub
Sub WBrw()
AcsVis WAcs
End Sub
Sub WCls()
On Error Resume Next
W.Close
Set W = Nothing
End Sub
Sub WDrp(TT)
DbDrpTbl W, TT
End Sub
Function PnmFfn$(A$)
PnmFfn = PnmPth(A) & PnmFn(A)
End Function
Function PnmPth$(A$)
PnmPth = PthEnsSfx(PnmVal(A & "Pth"))
End Function

Sub WReOpn()
WCls
WOpn
End Sub
Function RgRCC(A As Range, R, C1, C2) As Range
Set RgRCC = RgRCRC(A, R, C1, R, C2)
End Function

Sub ZZ_WtLnkFx()
AyDmp WtLnkFx(">UOM", IFxUOM)
End Sub
Sub LoSetFml(A As ListObject, ColNm$, Fml$)
A.ListColumns(ColNm).DataBodyRange.Formula = Fml
End Sub
Sub WtRen(Fmt$, ToT$, Optional ReOpnFst As Boolean)
DbtRen W, Fmt, ToT, ReOpnFst
End Sub


Sub WClr()
Exit Sub
Dim T, Tny$()
Tny = WTny: If Sz(Tny) = 0 Then Exit Sub
For Each T In Tny
    WDrp T
Next
End Sub
Function WTny() As String()
WTny = DbTny(W)
End Function
Function WStru$(Optional TT$)
If TT = "" Then
    WStru = DbStru(W)
Else
    WStru = DbttStru(W, TT)
End If
End Function
Function WAcs() As Access.Application
Set WAcs = ApnAcs(Apn)
End Function
Function WtFny(T$) As String()
WtFny = DbtFny(W, T)
End Function
Function WtStru$(T$)
WtStru = DbtStru(W, T)
End Function

Function WttStru$(TT)
WttStru = DbttStru(W, TT)
End Function
Function WFb$()
WFb = ApnWFb(Apn)
End Function

Sub WImp(T$, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then Stop
DbtImpMap W, T, LnkColStr, WhBExpr
End Sub

Sub FfnMov(Fm, ToFfn)
Fso.MoveFile Fm, ToFfn
End Sub
Sub CrtResTbl()
DbCrtResTbl CurrentDb
End Sub
Sub EnsResTbl()
DbEnsResTbl CurrentDb
End Sub
Sub DbEnsResTbl(A As Database)
If Not DbHasTbl(A, "Res") Then DbCrtResTbl A
End Sub

Sub DbCrtResTbl(A As Database)
DbtDrp A, "Res"
DoCmd.RunSQL "Create Table Res (ResNm Text(50), Att Attachment)"
End Sub
Function DbResExp$(A As Database, ResNm)
'Resnm is Tbl.Fld.Key  With Tbl-Dft and Fld-Dft as Res
'Export the res to tmpFfn and return tmpFfn
Dim O$
O = TmpFfn
DbResAttFld(A, ResNm).SaveToFile O
DbResExp = O
End Function
Function DbResAttFld(A As Database, ResNm) As Field2
End Function
Sub ResClr(A$)
DbResClr CurrentDb, A
End Sub
Sub DbResClr(A As Database, ResNm$)
A.Execute "Delete From Res where ResNm='" & ResNm & "'"
End Sub
Sub RsDmp(A As Recordset)
AyDmp RsCsvLy(A)
A.MoveFirst
End Sub
Sub RsDmpByFny0(A As Recordset, Fny0)
AyDmp RsCsvLyByFny0(A, Fny0)
A.MoveFirst
End Sub
Function AttRs(A$) As AttRs
AttRs = DbAttRs(CurrentDb, A)
End Function
Function AttFny() As String()
AttFny = ItrNy(DbFstAttRs(CurrentDb).AttRs.Fields)
End Function
Function AttRs_Exp$(A As AttRs, ToFfn$)
'Export the only File in {AttRs} {ToFfn}
Dim Fn$, Ext$, T$, F2 As DAO.Field2
With A.AttRs
    If FfnExt(CStr(!FileName)) <> FfnExt(ToFfn) Then Stop
    Set F2 = .Fields("FileData")
End With
F2.SaveToFile ToFfn
AttRs_Exp = ToFfn
End Function
Function DbAttExp$(A As Database, Att$, ToFfn$)
'Exporting the only file in Att
If Not DbAtt_HasOnlyOneFile(A, Att) Then Stop
DbAttExp = AttRs_Exp(DbAttRs(A, Att), ToFfn)
End Function
Function DbAttRs(A As Database, Att$) As AttRs
With DbAttRs
    Set .TblRs = A.OpenRecordset(FmtQQ("Select Att,FilTim,FilSz from Att where AttNm='?'", Att))
    If .TblRs.EOF Then
        A.Execute FmtQQ("Insert into Att (AttNm) values('?')", Att)
        Set .TblRs = A.OpenRecordset(FmtQQ("Select Att from Att where AttNm='?'", Att))
    End If
    Set .AttRs = .TblRs.Fields(0).Value
End With
End Function
Function DbFstAttRs(A As Database) As AttRs
With DbFstAttRs
    Set .TblRs = A.TableDefs("Att").OpenRecordset
    Set .AttRs = .TblRs.Fields("Att").Value
End With
End Function
Sub ZZ_DbAttExpFfn()
Dim T$
T = TmpFx
DbAttExpFfn CurrentDb, "Tp", "TaxRateAlert(Template).xlsm", T
Debug.Assert FfnIsExist(T)
Kill T
End Sub
Function DbAttExpFfn$(A As Database, Att$, AttFn$, ToFfn$)
Dim F2 As Field2, O$(), AttRs As AttRs
If FfnExt(AttFn) <> FfnExt(ToFfn) Then
    Stop
End If
If FfnIsExist(ToFfn) Then Stop
AttRs = DbAttRs(A, Att)
With AttRs
    With .AttRs
        .MoveFirst
        While Not .EOF
            If !FileName = AttFn Then
                Set F2 = !FileData
                F2.SaveToFile ToFfn
                DbAttExpFfn = ToFfn
                Exit Function
            End If
            .MoveNext
        Wend
        Push O, "Database          : " & A.Name
        Push O, "AttKey            : " & Att
        Push O, "Missing-AttFn     : " & AttFn
        Push O, "AttKey-File-Count : " & AttRs.AttRs.RecordCount
        PushAy O, AyAddPfx(RsSy(AttRs.AttRs, "FileName"), "Fn in AttKey      : ")
        Push O, "Att-Table in Database has AttKey, but no Fn-of-Ffn"
        AyBrw O
        Stop
        Exit Function
    End With
End With
If IsNothing(F2) Then Stop
F2.SaveToFile ToFfn
DbAttExpFfn = ToFfn
End Function
Sub AttClr(A$)
DbClrAtt CurrentDb, A
End Sub
Sub DbClrAtt(A As Database, Att$)
RsClr DbAttRs(A, Att).AttRs
End Sub
Sub RsClr(A As DAO.Recordset)
With A
    While Not .EOF
        .Delete
        .MoveNext
    Wend
End With
End Sub
Function AttExpFfn$(A$, AttFn$, ToFfn$)
AttExpFfn = DbAttExpFfn(CurrentDb, A, AttFn, ToFfn)
End Function

Function DbAttTblRs(A As Database, AttNm$) As DAO.Recordset
Set DbAttTblRs = A.OpenRecordset(FmtQQ("Select * from Att where AttNm='?'", AttNm))
End Function

Function DbAttFnAy(A As Database, Att$) As String()
Dim T As DAO.Recordset ' AttTblRs
Dim F As DAO.Recordset ' AttFldRs
Set T = DbAttTblRs(A, Att)
Set F = T.Fields("Att").Value
DbAttFnAy = RsSy(F, "FileName")
End Function

Function AttFnAy(A$) As String()
AttFnAy = DbAttFnAy(CurrentDb, "AA")
End Function

Function ZZ_AttFnAy()
D AttFnAy("AA")
End Function
Sub ZZ_AttImp()
Dim T$
T = TmpFt
StrWrt "sdfdf", T
AttImp "AA", T
Kill T
'T = TmpFt
'AttExpFfn "AA", T
'FtBrw T
End Sub

Function RsMovFst(A As DAO.Recordset) As DAO.Recordset
A.MoveFirst
Set RsMovFst = A
End Function
Function AttFfn$(A$)
'Return Fst-Ffn-of-Att-A
AttFfn = RsMovFst(AttRs(A).AttRs)!FileName
End Function
Function AttHasOnlyOneFile(A$) As Boolean
AttHasOnlyOneFile = DbAtt_HasOnlyOneFile(CurrentDb, A)
End Function
Function DbAtt_HasOnlyOneFile(A As Database, Att$) As Boolean
Debug.Print "DbAtt_HasOnlyFile: " & DbAttRs(A, Att).AttRs.RecordCount
DbAtt_HasOnlyOneFile = DbAttRs(A, Att).AttRs.RecordCount = 1
End Function

Function AttExp$(A$, ToFfn$)
'Exporting the only file in Att
AttExp = DbAttExp(CurrentDb, A, ToFfn)
Debug.Print "-----"
Debug.Print "AttExp"
Debug.Print "Att   : "; A
Debug.Print "ToFfn : "; ToFfn
Debug.Print "Att is: Export to ToFfn"
End Function
Sub AttImp(A$, FmFfn$)
'Exporting the only file in Att
DbAttImp CurrentDb, A, FmFfn
End Sub
Sub DbAttImp(A As Database, Att$, FmFfn$)
AttRs_Imp DbAttRs(A, Att), FmFfn
End Sub
Function AttFstFn$(A$)
AttFstFn = DbAtt_FstFn(CurrentDb, A)
End Function
Function DbAtt_FstFn(A As Database, Att$)
DbAtt_FstFn = DbAttRs(A, Att).AttRs!FileName
End Function
Function RsHasFldV(A As DAO.Recordset, F$, V) As Boolean
With A
    .MoveFirst
    While Not .EOF
        If .Fields(F) = V Then RsHasFldV = True: Exit Function
    Wend
End With
End Function
Sub AttRs_Imp(A As AttRs, Ffn$)
Dim F2 As Field2
Const Trc As Boolean = True
Dim S&, T As Date
S = FfnSz(Ffn)
T = FfnTim(Ffn)
If Trc Then
    Debug.Print "----------"
    Debug.Print "AttRs_Imp:"
    Debug.Print "Att       : "; A.TblRs.Name
    Debug.Print "Given-File: "; Ffn
    Debug.Print "Given-File-Sz : "; S
    Debug.Print "Given-File-Tim: "; T
End If
With A
    .TblRs.Edit
    With .AttRs
        If RsHasFldV(A.AttRs, "FileName", FfnFn(Ffn)) Then
            If Trc Then Debug.Print "Given-Ffn is : Found in Att, it is replaced"
            .Edit
        Else
            If Trc Then Debug.Print "Given-Ffn is : Not found in Att, it is added"
            .AddNew
        End If
        Set F2 = !FileData
        F2.LoadFromFile Ffn
        .Update
    End With
    .TblRs.Fields!FilTim = FfnTim(Ffn)
    .TblRs.Fields!FilSz = FfnSz(Ffn)
    .TblRs.Update
End With
End Sub

Function FxDaoCnStr$(A)
'Excel 8.0;HDR=YES;IMEX=2;DATABASE=D:\Data\MyDoc\Development\ISS\Imports\PO\PUR904 (On-Line).xls;TABLE='PUR904 (On-Line)'
'INTO [Excel 8.0;HDR=YES;IMEX=2;DATABASE={0}].{1} FROM {2}"
'Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=C:\Users\sium\Desktop\TaxRate\sales text.xlsx;TABLE=Sheet1$
Dim O$
Select Case FfnExt(A)
Case ".xlsx":: O = "Excel 12.0 Xml;HDR=YES;IMEX=2;ACCDB=YES;DATABASE=" & A & ";"
Case ".xls": O = "Excel 8.0;HDR=YES;IMEX=2;DATABASE=" & A & ";"
Case Else: Stop
End Select
FxDaoCnStr = O
End Function
Sub ZZ_FbWb_zExpOupTbl()
Dim W As Workbook
Set W = FbWb_zExpOupTbl(WFb)
WbVis W
Stop
W.Close False
Set W = Nothing
End Sub
Function AyWhHasPfx(A, Pfx$) As String()
AyWhHasPfx = AyWhPredXP(A, "HasPfx", Pfx)
End Function
Sub ZZ_FbOupTny()
D FbOupTny(WFb)
End Sub
Function FbOupTny(A$) As String()
FbOupTny = AyWhHasPfx(FbTny(A), "@")
End Function

Sub AyRunABX(Ay, ABX$, A, B)
If Sz(Ay) = 0 Then Exit Sub
Dim X
For Each X In Ay
    Run ABX, A, B, X
Next
End Sub

Sub FbWrtFx_zForExpOupTb(A$, Fx$)
FbWb_zExpOupTbl(A).SaveAs Fx
End Sub
Function WcAddWs(A As WorkbookConnection) As Worksheet
Dim Wb As Workbook, Ws As Worksheet, Lo As ListObject, Qt As QueryTable
Set Wb = A.Parent
Set Ws = WbAddWs(Wb, A.Name)
Ws.Name = A.Name
Set Lo = Ws.ListObjects.Add(SourceType:=0, Source:=A.OLEDBConnection.Connection, Destination:=WsA1(Ws))
With Lo.QueryTable
    .CommandType = xlCmdTable
    .CommandText = A.Name
    .RowNumbers = False
    .FillAdjacentFormulas = False
    .PreserveFormatting = True
    .RefreshOnFileOpen = False
    .BackgroundQuery = True
    .RefreshStyle = xlInsertDeleteCells
    .SavePassword = False
    .SaveData = True
    .AdjustColumnWidth = True
    .RefreshPeriod = 0
    .PreserveColumnInfo = True
    .ListObject.DisplayName = TnLoNm(A.Name)
    .Refresh BackgroundQuery:=False
End With
Set WcAddWs = Ws
End Function

Function FbWb_zExpOupTbl(A$) As Workbook
Dim O As Workbook
Set O = NewWb
AyRunABX FbOupTny(A), "WbAddWc", O, A
ItrDo O.Connections, "WcAddWs"
WbRfh O
Set FbWb_zExpOupTbl = O
End Function
Sub PushObj_zNonNothing(Oy, Obj)
If IsNothing(Obj) Then Exit Sub
PushObj Oy, Obj
End Sub
Function WbWcAy_zOle(A As Workbook) As OLEDBConnection()
Dim O() As OLEDBConnection, Wc As WorkbookConnection
For Each Wc In A.Connections
    PushObj_zNonNothing O, Wc.OLEDBConnection
Next
WbWcAy_zOle = O
End Function
Function WbWcSy_zOle(A As Workbook) As String()
WbWcSy_zOle = OyPrpSy(WbWcAy_zOle(A), "Connection")
End Function
Sub ZZ_WbWcSy()
D WbWcSy_zOle(FxWb(TpFx))
End Sub
Function WbWcNy(A As Workbook) As String()
WbWcNy = ItrNy(A.Connections)
End Function
Sub WbAddWc(A As Workbook, Fb$, Nm$)
A.Connections.Add2 Nm, Nm, FbWcStr(Fb), Nm, XlCmdType.xlCmdTable
End Sub

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function
Function AyKeepLasN(A, N)
Dim O, J&, I&, U&, Fm&, NewU&
U = UB(A)
If U < N Then AyKeepLasN = A: Exit Function
O = A
Fm = U - N + 1
NewU = N - 1
For J = Fm To U
    Asg O(J), O(I)
    I = I + 1
Next
ReDim Preserve O(NewU)
AyKeepLasN = O
End Function
Sub ZZ_LinesKeepLasN()
Dim Ay$(), A$, J%
For J = 0 To 9
Push Ay, "Line " & J
Next
A = Join(Ay, vbCrLf)
'Debug.Print fLasN(A, 3)
End Sub
Function LinesKeepLasN$(A$, N%)
Dim Ay$()
Ay = SplitCrLf(A)
LinesKeepLasN = JnCrLf(AyKeepLasN(Ay, N))
End Function
Function FbDaoCn(A) As DAO.Connection
Set FbDaoCn = DBEngine.OpenConnection(A)
End Function
Function CvCtl(A) As Access.Control
Set CvCtl = A
End Function
Function CvBtn(A) As Access.CommandButton
Set CvBtn = A
End Function
Function IsBtn(A) As Boolean
IsBtn = TypeName(A) = "CommandButton"
End Function
Function IsTgl(A) As Boolean
IsTgl = TypeName(A) = "ToggleButton"
End Function
Function CvTgl(A) As Access.ToggleButton
Set CvTgl = A
End Function
Sub CmdTurnOffTabStop(AcsCtl)
Dim A As Access.Control
Set A = AcsCtl
If Not HasPfx(A.Name, "Cmd") Then Exit Sub
Select Case True
Case IsBtn(A): CvBtn(A).TabStop = False
Case IsTgl(A): CvTgl(A).TabStop = False
End Select
End Sub
Sub FrmSetCmdNotTabStop(A As Access.Form)
ItrDo A.Controls, "CmdTurnOffTabStop"
End Sub
Function FxAdoCnStr$(A)
FxAdoCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.16.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", A)
End Function
Function FxOleCnStr$(A)
FxOleCnStr = "OLEDb;" & FxAdoCnStr(A)
End Function
Function FbOleCnStr$(A)
FbOleCnStr = "OLEDb;" & FbAdoCnStr(A)
End Function

Function FbAdoCnStr$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
FbAdoCnStr = FmtQQ(C, A)
End Function
Function AdoCnStr_Cn(A) As adodb.Connection
Dim O As New adodb.Connection
O.Open A
Set AdoCnStr_Cn = O
End Function
Function FxCn(A) As adodb.Connection
Set FxCn = AdoCnStr_Cn(FxAdoCnStr(A))
End Function
Function FbCn(A) As adodb.Connection
Set FbCn = AdoCnStr_Cn(FbAdoCnStr(A))
End Function

Function FxCat(A) As Catalog
Set FxCat = CnCat(FxCn(A))
End Function

Function CnCat(A As adodb.Connection) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = A
Set CnCat = O
End Function

Function FbTny(A$) As String()
Dim Db As Database
Set Db = FbDb(A)
FbTny = DbTny(Db)
Db.Close
'FbTny = CvSy(AyWhPredXPNot(CatTny(FbCat(A)), "HasPfx", "MSys"))
End Function

Function AyCln(A)
Dim O
O = A
Erase O
AyCln = O
End Function
Function AyWhPredXPNot(A, PredXP$, P)
If Sz(A) = 0 Then AyWhPredXPNot = AyCln(A): Exit Function
Dim O, X
O = AyCln(A)
For Each X In A
    If Not Run(PredXP, X, P) Then
        Push O, X
    End If
Next
AyWhPredXPNot = O
End Function
Function AyWhPredXP(A, PredXP$, P)
If Sz(A) = 0 Then AyWhPredXP = AyCln(A): Exit Function
Dim O, X
O = AyCln(A)
For Each X In A
    If Run(PredXP, X, P) Then
        Push O, X
    End If
Next
AyWhPredXP = O
End Function
Function FbCat(A) As Catalog
Set FbCat = CnCat(FbCn(A))
End Function
Function CatTny(A As Catalog) As String()
CatTny = ItrNy(A.Tables)
End Function
Function FxWsNy(A) As String()
FxWsNy = CatTny(FxCat(A))
End Function
Function FxHasWs(A, WsNm$) As Boolean
FxHasWs = AyHas(FxWsNy(A), WsNm)
End Function

Sub DbtImpTbl(A As Database, TT)
Dim Tny$(), J%, S$
Tny = CvNy(TT)
For J = 0 To UB(Tny)
    DbtDrp A, "#I" & Tny(J)
    S = FmtQQ("Select * into [#I?] from [?]", Tny(J), Tny(J))
    A.Execute S
Next
End Sub
Function LnkColStr_Ly(A$) As String()
Dim A1$(), A2$(), Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
A1 = LnkColAy_Ny(Ay)
A2 = AyAlignL(AyQuoteSqBkt(LnkColAy_ExtNy(Ay)))
Dim J%, O$()
For J = 0 To UB(A1)
    Push O, A2(J) & "  " & A1(J)
Next
LnkColStr_Ly = O
End Function
Function AyLasEle(A)
Asg A(UB(A)), AyLasEle
End Function

Function AscIsDig(A%) As Boolean
AscIsDig = &H30 <= A And A <= &H39
End Function

Property Get LnkCol(Nm$, Ty As DAO.DataTypeEnum, Extnm$) As LnkCol
Dim O As New LnkCol
Set LnkCol = O.Init(Nm, Ty, Extnm)
End Property

Function LnkColStr_LnkColAy(A) As LnkCol()
Dim Emp() As LnkCol, Ay$()
Ay = SplitVBar(A): If Sz(Ay) = 0 Then Stop
LnkColStr_LnkColAy = AyMapInto(Ay, "LinLnkCol", Emp)
End Function

Function SplitVBar(A) As String()
SplitVBar = Split(A, "|")
End Function

Function RmvSqBkt$(A)
If IsSqBktQuoted(A) Then
    RmvSqBkt = RmvFstLasChr(A)
Else
    RmvSqBkt = A
End If
End Function
Sub ZZ_LinLnkCol()
Dim A$, Act As LnkCol, Exp As LnkCol
A = "AA Txt XX"
Exp = LnkCol("AA", dbText, "AA")
GoSub Tst
Exit Sub
Tst:
Act = LinLnkCol(A)
Debug.Assert LnkColIsEq(Act, Exp)
Return
End Sub
Function LnkColIsEq(A As LnkCol, B As LnkCol) As Boolean
With A
    If .Extnm <> B.Extnm Then Exit Function
    If .Ty <> B.Ty Then Exit Function
    If .Nm <> B.Nm Then Exit Function
End With
LnkColIsEq = True
End Function
Function LinLnkCol(A) As LnkCol
Dim Nm$, ShtTy$, Extnm$, Ty As DAO.DataTypeEnum
LinTTRstAsg A, Nm, ShtTy, Extnm
Extnm = RmvSqBkt(Extnm)
Ty = DaoShtTy_Ty(ShtTy)
Set LinLnkCol = LnkCol(Nm, Ty, IIf(Extnm = "", Nm, Extnm))
End Function
Function RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Function
Function DbtCnStr$(A As Database, T)
DbtCnStr = A.TableDefs(T).Connect
End Function
Sub DbtImpMap(A As Database, T$, LnkColStr$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "FstChr of T must be >"
    Stop
End If
'Assume [>?] T exist
'Create [#I?] T
Dim S$
S = LnkColStr_ImpSql(LnkColStr, T, WhBExpr)
DbtDrp A, "#I" & Mid(T, 2)
A.Execute S
End Sub

Function LnkColStr_ImpSql$(A$, T$, Optional WhBExpr$)
Dim Ay() As LnkCol
Ay = LnkColStr_LnkColAy(A)
LnkColStr_ImpSql = LnkColAy_ImpSql(Ay, T, WhBExpr)
End Function

Function IsSqBktQuoted(A) As Boolean
If FstChr(A) <> "[" Then Exit Function
If LasChr(A) <> "]" Then Exit Function
IsSqBktQuoted = True
End Function

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function LasChr$(A)
LasChr = Right(A, 1)
End Function
Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(CvNy(Fny0), Dry)
End Property
Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Dim O$(), J%, U&
U = UB(Av)
ReDim O(U)
For J = 0 To U
    O(J) = Av(J)
Next
ApSy = O
End Function
Function DbtHasFld(A As Database, T$, F$) As Boolean
DbtHasFld = ItrHasNm(A.TableDefs(T).Fields, F)
End Function
Sub ZZ_SampleLo()
LoVis SampleLo
End Sub
Function SampleLo() As ListObject
Set SampleLo = DrsLo(SampleDrs, NewA1, "T_Sample")
End Function
Function DrsLo(A As Drs, At As Range, Optional LoNm$) As ListObject
Set DrsLo = RgLo(SqRg(DrsSq(A), At), LoNm)
End Function
Function SqRg(A, At As Range) As Range
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqRg = O
End Function

Function SampleDrs() As Drs
Set SampleDrs = Drs("A B C D E F", SampleDry)
End Function
Function SampleDry() As Variant()
Dim O(), Dr(), I%, J%
For J = 0 To 9
    ReDim Dr(5)
    For I = 0 To 5
        Dr(I) = J * 100 + I
    Next
    Push O, Dr
Next
SampleDry = O
End Function
Function AyIdx&(A, Itm)
AyIdx = AyIdxFm(A, Itm, 0)
End Function
Function AyIdxFm&(A, Itm, Fm&)
Dim O&
For O = Fm To UB(A)
    If A(O) = Itm Then AyIdxFm = O: Exit Function
Next
AyIdxFm = -1
End Function
Sub ZZ_AyHasAyInSeq()
Dim A, B
A = Array(1, 2, 3, 4, 5, 6, 7, 8)
B = Array(2, 4, 6)
Debug.Assert AyHasAyInSeq(A, B) = True

End Sub
Function AyHasAyInSeq(A, B) As Boolean
Dim BItm, Ix&
If Sz(B) = 0 Then Stop
For Each BItm In B
    Ix = AyIdxFm(A, BItm, Ix)
    If Ix = -1 Then Exit Function
    Ix = Ix + 1
Next
AyHasAyInSeq = True
End Function
Sub LoSetOutLin(A As ListObject, L1Ny0)
Dim L1Ny$(), LFny$()
LFny = LoFny(A)
L1Ny = CvNy(L1Ny0)
If AyHasAyInSeq(LFny, L1Ny) Then Stop
Dim C As ListColumn
For Each C In A.ListColumns
    If Not AyHas(L1Ny, C.Name) Then
        C.Range.EntireColumn.OutlineLevel = 2
    End If
Next
End Sub
Function SqRplLo(A, Lo As ListObject) As ListObject
Dim LoNm$, At As Range
LoNm = Lo.Name
Set At = Lo.Range
Lo.Delete
Set SqRplLo = RgLo(SqRg(A, At), LoNm)
End Function
Function SqAt_Lo(A, At As Range, Optional LoNm$) As ListObject
Set SqAt_Lo = RgLo(SqRg(A, At), LoNm)
End Function
Function WbLoAy(Tp As Workbook) As ListObject()
Dim Ws As Worksheet, O() As ListObject, I
For Each Ws In Tp.Sheets
    OyPushItr O, Ws.ListObjects
Next
WbLoAy = O
End Function
Sub OyPushItr(Oy, Itr)
Dim I
For Each I In Itr
    PushObj Oy, I
Next
End Sub
Function AscIsUCase(A%) As Boolean
AscIsUCase = 65 <= A And A <= 90
End Function
Function AscIsLCase(A%) As Boolean
AscIsLCase = 97 <= A And A <= 122
End Function
Function AscIsLetter(A%) As Boolean
AscIsLetter = True
If AscIsUCase(A) Then Exit Function
If AscIsLCase(A) Then Exit Function
AscIsLetter = False
End Function
Function RmvFstNonLetter$(A)
If AscIsLetter(Asc(A)) Then
    RmvFstNonLetter = A
Else
    RmvFstNonLetter = RmvFstChr(A)
End If
End Function
Function AyRmvFstNonLetter(A) As String()
AyRmvFstNonLetter = AyMapSy(A, "RmvFstNonLetter")
End Function
Function DbtNewWb(A As Database, TT) As Workbook

End Function

Function DbtRplLo(A As Database, T$, Lo As ListObject, Optional ReSeqSpec$) As ListObject
Set DbtRplLo = SqRplLo(DbtSq(A, T, ReSeqSpec), Lo)
End Function
Sub ZZ_LoKeepFstCol()
LoKeepFstCol LoVis(SampleLo)
End Sub
Sub LoKeepFstCol(A As ListObject)
Dim J%
For J = A.ListColumns.Count To 2 Step -1
    A.ListColumns(J).Delete
Next
End Sub
Function WbLo(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If WsHasLo(Ws, LoNm) Then Set WbLo = Ws.ListObjects(LoNm): Exit Function
Next
End Function
Function WsHasLo(A As Worksheet, LoNm$) As Boolean
WsHasLo = ItrHasNm(A.ListObjects, LoNm)
End Function
Sub LoKeepFstRow(A As ListObject)
Dim J%
For J = A.ListRows.Count To 2 Step -1
    A.ListRows(J).Delete
Next
End Sub
Function DbDrpTbl(A As Database, TT)
AyDoPX CvNy(TT), "DbtDrp", A
End Function
Sub SavRec()
DoCmd.RunCommand acCmdSaveRecord
End Sub

Sub AyDoPX(A, PXFunNm$, P)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run PXFunNm, P, I
Next
End Sub
Function DbqRs(A As Database, Sql) As DAO.Recordset
Set DbqRs = A.OpenRecordset(Sql)
End Function
Function Acs() As Access.Application
Static X As Boolean, Y As Access.Application
On Error GoTo X
If X Then
    Set Y = New Access.Application
    X = True
End If
If Y.Application.Name = "Microsoft Access" Then
    Set Acs = Y
    Exit Function
End If
X:
    Set Y = New Access.Application
    Debug.Print "Acs: New Acs instance is crreated."
Set Acs = Y
End Function

Sub AcsVis(A As Access.Application)
If Not A.Visible Then A.Visible = True
End Sub

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function
Function AyAddPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), U&, J&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
AyAddPfx = O
End Function
Function IsObjAy(A) As Boolean
IsObjAy = VarType(A) = vbArray + vbObject
End Function
Function AyRmvEleAt(A, Optional At&)
Dim O, J&, U&
U = UB(A)
O = A
Select Case True
Case U = 0
    Erase O
    AyRmvEleAt = O
    Exit Function
Case IsObjAy(A)
    For J = At To U - 1
        Set O(J) = O(J + 1)
    Next
Case Else
    For J = At To U - 1
        O(J) = O(J + 1)
    Next
End Select
ReDim Preserve O(U - 1)
AyRmvEleAt = O
End Function
Function AbIsEq(A, B) As Boolean
If VarType(A) <> VarType(B) Then Exit Function
Select Case True
Case IsObject(A): AbIsEq = ObjPtr(A) = ObjPtr(B)
Case IsArray(A): AbIsEq = AyIsEq(A, B)
Case Else: AbIsEq = A = B
End Select
End Function
Private Sub ZZZ_AyShift()
Dim Ay(), Exp, Act, ExpAyAft()
Ay = Array(1, 2, 3, 4)
Exp = 1
ExpAyAft = Array(2, 3, 4)
GoSub Tst
Exit Sub
Tst:
Act = AyShift(Ay)
Debug.Assert AbIsEq(Exp, Act)
Debug.Assert AyIsEq(Ay, ExpAyAft)
Return
End Sub


Function AyShift(Ay)
AyShift = Ay(0)
Ay = AyRmvEleAt(Ay)
End Function
Private Sub ZZZ_PfxSsl_Sy()
Dim A$, Exp$()
A = "A B C D"
Exp = SslSy("AB AC AD")
GoSub Tst
Exit Sub
Tst:
Dim Act$()
Act = PfxSsl_Sy(A)
Debug.Assert AyIsEq(Act, Exp)
Return
End Sub
Function ItrFstPrpEq(A, PrpNm$, V)
Dim I, OP
For Each I In A
    OP = ObjPrp(I, PrpNm)
    If OP = V Then Asg I, ItrFstPrpEq: Exit Function
Next
Debug.Print PrpNm, V
For Each I In A
    Debug.Print ObjPrp(I, PrpNm)
Next
Stop
End Function
Function ObjPrp(A, PrpNm$)
On Error GoTo X
Dim V
V = CallByName(A, PrpNm, VbGet)
Asg V, ObjPrp
Exit Function
X:
Debug.Print "ObjPrp: " & Err.Description
End Function
Function ItrPrpSy(A, PrpNm$) As String()
ItrPrpSy = ItrPrpInto(A, PrpNm, EmpSy)
End Function
Function ItrPrpInto(A, PrpNm$, OInto)
Dim O, I
O = OInto
Erase O
For Each I In A
    Push O, ObjPrp(I, PrpNm)
Next
ItrPrpInto = O
End Function
Function WbWsCdNy(A As Workbook) As String()
WbWsCdNy = ItrPrpSy(A.Sheets, "CodeName")
End Function
Function FxWsCdNy(A) As String()
Dim Wb As Workbook
Set Wb = FxWb(A)
FxWsCdNy = WbWsCdNy(Wb)
Wb.Close False
End Function
Function PfxSsl_Sy(A) As String()
Dim Ay$(), Pfx$
Ay = SslSy(A)
Pfx = AyShift(Ay)
PfxSsl_Sy = AyAddPfx(Ay, Pfx)
End Function
Function ApnWAcs(A$)
Dim O As Access.Application
AcsOpn O, ApnWFb(A)
Set ApnWAcs = O
End Function
Function ApnAcs(A$) As Access.Application
AcsOpn Acs, ApnWFb(A)
Set ApnAcs = Acs
End Function
Sub AcsOpn(A As Access.Application, Fb$)
Select Case True
Case IsNothing(A.CurrentDb)
    A.OpenCurrentDatabase Fb
Case A.CurrentDb.Name = Fb
Case Else
    A.CurrentDb.Close
    A.OpenCurrentDatabase Fb
End Select
End Sub
Sub ApnBrwWDb(A$)
Dim Fb$
Fb = ApnWFb(A)
AcsOpn Acs, Fb
AcsVis Acs
End Sub
Sub FbEns(A$)
If FfnIsExist(A) Then Exit Sub
FbCrt A
End Sub
Sub FbCrt(A$)
DBEngine.CreateDatabase A, dbLangGeneral
End Sub
Sub FxRfhWbWcStr(A, Fb$)
WbRfhCnStr(FxWb(A), Fb).Close True
End Sub
Function WbRfhCnStr(A As Workbook, Fb$) As Workbook
ItrDoXP A.Connections, "WcRfhCnStr", FbWcStr(Fb)
Set WbRfhCnStr = A
End Function
Sub FbOpn(A)
Acs.OpenCurrentDatabase A
AcsVis Acs
End Sub
Function FbDb(A) As Database
Set FbDb = DBEngine.OpenDatabase(A)
End Function

Function ApnWFb$(A$)
ApnWFb = ApnWPth(A) & "Wrk.accdb"
End Function

Function ApnWPth$(A$)
Dim P$
P = TmpHom & A & "\"
PthEns P
ApnWPth = P
End Function
Function DbIsOk(A As Database) As Boolean
On Error GoTo X
DbIsOk = IsStr(A.Name)
Exit Function
X:
End Function
Function WsC(A As Worksheet, C) As Range
Dim R As Range
Set R = A.Columns(C)
Set WsC = R.EntireColumn
End Function

Function ApnWDb(A$) As Database
Static X As Boolean, Y As Database
If Not X Then
    X = True
    FbEns ApnWFb(A)
    Set Y = FbDb(ApnWFb(A))
End If
If Not DbIsOk(Y) Then Set Y = FbDb(ApnWFb(A))
Set ApnWDb = Y
End Function
Function DbqAny(A As Database, Sql) As Boolean
DbqAny = RsAny(DbqRs(A, Sql))
End Function
Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type in (1,6)", T))
End Function
Function AyWdt%(A)
Dim O%, J&
For J = 0 To UB(A)
    O = Max(O, Len(A(J)))
Next
AyWdt = O
End Function
Function TTStru$(TT)
TTStru = DbttStru(CurrentDb, TT)
End Function
Function TblStru$(T$)
TblStru = DbtStru(CurrentDb, T)
End Function

Function QTbl$(T$, Optional WhBExpr$)
QTbl = "Select *" & PFm(T) & PWh(WhBExpr)
End Function
Function FbtFny(A$, T$) As String()
FbtFny = RsFny(DbqRs(FbDb(A), QTbl(T)))
End Function
Function Max(A, B)
If A > B Then
    Max = A
Else
    Max = B
End If
End Function
Function Min(A, B)
If A > B Then
    Min = B
Else
    Min = A
End If
End Function

Function CvNy(Ny0) As String()
Select Case True
Case IsMissing(Ny0)
Case IsStr(Ny0): CvNy = SslSy(Ny0)
Case IsSy(Ny0): CvNy = Ny0
Case IsArray(Ny0): CvNy = AySy(Ny0)
Case Else: Stop
End Select
End Function
Function AySy(A) As String()
If Sz(A) = 0 Then Exit Function
AySy = ItrAy(A, EmpSy)
End Function
Function EmpSy() As String()
End Function
Function EmpAy() As Variant()
End Function
Sub TpMinLo()
Dim O As Workbook
Set O = TpWb
WbMinLo O
O.Save
WbVis O
End Sub

Function TpIdxWs() As Worksheet
Set TpIdxWs = WbWsCd(TpWb, "WsIdx")
End Function
Function TpWsCdNy() As String()
TpWsCdNy = FxWsCdNy(TpFx)
End Function

Function TpWcSy() As String()
Dim W As Workbook, X As Excel.Application
Set X = New Excel.Application
Set W = X.Workbooks.Open(TpFx)
TpWcSy = WbWcSy_zOle(W)
W.Close False
Set W = Nothing
X.Quit
Set X = Nothing
End Function

Function TpFx$()
TpFx = TpPth & Apn & "(Template).xlsm"
End Function
Sub TpOpn()
FxOpn TpFx
End Sub
Function TpWb() As Workbook
Set TpWb = FxWb(TpFx)
End Function

Function ItrAy(A, OInto)
Dim O, I
O = OInto
Erase O
For Each I In A
    Push O, I
Next
ItrAy = O
End Function

Function OupPth$()
Dim A$
A = CurDbPth & "Output\"
PthEns A
OupPth = A
End Function
Function OupPth_zPm$()
OupPth_zPm = PnmVal("OupPth")
End Function
Function YYYYMMDD_IsVdt(A) As Boolean
On Error Resume Next
YYYYMMDD_IsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function
Function TpPth$()
TpPth = PthEns(CurDbPth & "Template\")
End Function
Function FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
FfnPth = Left(A, P)
End Function
Private Function ErzFws__2(Fx$, WsNm$, ColNy$()) As String()

End Function
Private Function ErzFws__3(Fx$, WsNm$, ColNy$(), DtaTyAy() As DAO.DataTypeEnum) As String()

End Function
Sub ZZ_ErAyzFxWsMissingCol()
'" [Material]             As Sku," & _
'" [Plant]                As Whs," & _
'" [Storage Location]     As Loc," & _
'" [Batch]                As BchNo," & _
'" [Unrestricted]         As OH " & _

End Sub
Function TblF_Ty(T, F) As DAO.DataTypeEnum

End Function
Function TblErAyzCol(T$, ColNy$(), DtaTyAy() As DAO.DataTypeEnum, Optional AddTblLinMsg As Boolean) As String()
Dim Fny$(), F, Fny1$(), Fny2$()
Fny = TblFny(T)
For Each F In ColNy
    If AyHas(Fny, F) Then
        Push F, Fny1
    Else
        Push F, Fny2
    End If
Next
Dim O$()
If Sz(Fny2) > 0 Then
    Dim J%
    For J = 0 To UB(ColNy)
        If AyHas(Fny2, ColNy(J)) Then
            If TblF_Ty(T, ColNy(J)) <> DtaTyAy(J) Then
                Push O, "Column [?] has unexpected DataType[?].  It is expected to be [?]"
            End If
        End If
    Next
End If
If AddTblLinMsg Then
    Push O, ""
    
End If
End Function
Function ErzFfnNotExist(A$) As String()
Dim O$(), M$
If Not FfnIsExist(A$) Then
    Push O, A
    M = "Above file not exist"
    Push O, M
    Push O, String(Len(M), "-")
End If
ErzFfnNotExist = O
End Function
Function ErzThen(ParamArray ErFunNmAp()) As String()
Dim Av(), O$(), I
Av = ErFunNmAp
For Each I In Av
    O = Run(I)
    If Sz(O) > 0 Then
        ErzThen = O
    End If
Next
End Function
Function UnderLin$(A$)
UnderLin = String(Len(A), "-")
End Function
Function UnderLinDbl$(A$)
UnderLinDbl = String(Len(A), "=")
End Function
Function ErzFxWs(A$, WsNm$) As String()
'ErThen "ErzFfnNotExist ErzFxHasNoWs"
Dim O$()
O = ErzFfnNotExist(A)
If Sz(O) > 0 Then
    ErzFxWs = O
    Exit Function
End If

'B = ErzFxWs__1(A, WsNm)
If Sz(A) > 0 Then
'    ErAyzFxWs = A
    Exit Function
End If


If Not FfnIsExist(A) Then
    Push O, A
    Push O, "Above Excel file not found"
    Push O, "--------------------------"
    'ErAyzFxWsLnk = O
    Exit Function
End If
Dim B$
'B = FxWs_LnkErMsg(A, WsNm)
If B <> "" Then
    Push O, "Excel File: " & A
    Push O, "Worksheet : " & WsNm
    Push O, "System Msg: " & B
    Push O, "Above Excel file & Worksheet cannot be linked to Access"
    Push O, "-------------------------------------------------------"
    'ErAyzFxWsLnk = O
    Exit Function
End If
On Error GoTo X
TblLnkFx "#", CStr(A), WsNm
TblDrp "#"
Exit Function
X:
'FxWs_LnkErMsg = Err.Description

'A = ErAyzFxWsMissingCol(
End Function
Function CurDbPth$()
CurDbPth = FfnPth(CurrentDb.Name)
End Function
Property Get PnmVal$(Pnm$)
PnmVal = CurrentDb.TableDefs("Prm").OpenRecordset.Fields(Pnm).Value
End Property
Property Let PnmVal(Pnm$, V$)
Stop
'Should not use
With CurrentDb.TableDefs("Prm").OpenRecordset
    .Edit
    .Fields(Pnm).Value = V
    .Update
End With
End Property

Function FldsFny(A As DAO.Fields) As String()
FldsFny = ItrNy(A)
End Function
Sub PthBrw(A)
Shell FmtQQ("Explorer ""?""", A), vbMaximizedFocus
End Sub
Function PthEnsSfx$(A)
If Right(A, 1) <> "\" Then
    PthEnsSfx = A & "\"
Else
    PthEnsSfx = A
End If
End Function
Function ItrNy(A) As String()
Dim O$(), I
For Each I In A
    Push O, I.Name
Next
ItrNy = O
End Function
Sub Push(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub
Sub PushObj(O, M)
Dim N&
N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub
Private Sub ZZ_PthFxAy()
Dim A$()
A = PthFxAy(CurDir)
AyDmp A
End Sub

Function DteIsVdt(A$) As Boolean
On Error Resume Next
DteIsVdt = Format(CDate(A), "YYYY-MM-DD") = A
End Function
Private Sub ZZ_TblFny()
AyDmp TblFny(">KE24")
End Sub
Function RsSy(A As DAO.Recordset, Optional FldNm$) As String()
Dim O$(), Ix
Ix = IIf(FldNm = "", 0, FldNm)
With A
    .MoveFirst
    While Not .EOF
        Push O, .Fields(Ix).Value
        .MoveNext
    Wend
End With
RsSy = O
End Function
Sub ZZ_SqlFny()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyDmp SqlFny(S)
End Sub
Function SqlFny(A) As String()
SqlFny = RsFny(SqlRs(A))
End Function
Sub ZZ_SqlRs()
Const S$ = "SELECT qSku.*" & _
" FROM [N:\SAPAccessReports\DutyPrepay5\DutyPrepay5 (With Import).accdb].[qSku] AS qSku;"
AyBrw RsCsvLy(SqlRs(S))
End Sub

Function SqlRs(A) As DAO.Recordset
Set SqlRs = CurrentDb.OpenRecordset(A)
End Function
Private Sub ZZ_SqlSy()
AyDmp SqlSy("Select Distinct UOR from [>Imp]")
End Sub
Function SqpzInBExpr$(Ay, FldNm$, Optional WithQuote As Boolean)
Const C$ = "[?] in (?)"
Dim B$
    If WithQuote Then
        B = JnComma(AyQuoteSng(Ay))
    Else
        B = JnComma(Ay)
    End If
SqpzInBExpr = FmtQQ(C, FldNm, B)
End Function
Function SqlSy(A) As String()
SqlSy = DbqSy(CurrentDb, A)
End Function
Function AyAdd(A, B)
Dim O
O = A
PushAy O, B
AyAdd = O
End Function
Sub ZZ_DbtWhDupKey()
TblDrp "#A #B"
DoCmd.RunSQL "Select Distinct Sku,BchNo,CLng(Rate) as RateRnd into [#A] from ZZ_DbtUpdSeq"
DbtWhDupKey CurrentDb, "#A", "Sku BchNo", "#B"
TTBrw "#B"
Stop
TblDrp "#B"
End Sub
Sub TTWbBrw(TT, Optional UseWc As Boolean)
WbVis TTWb(TT, UseWc)
End Sub
Sub TblBrw(T)
DoCmd.OpenTable T
End Sub
Function CvTT(A) As String()
CvTT = CvNy(A)
End Function

Sub TTBrw(TT)
'OFunAyDo DoCmd, "OpenTable", CvTT(TT)
End Sub

Sub DbtWhDupKey(A As Database, T$, KK, TarTbl$)
Dim Ky$(), K$, Jn$, Tmp$, J%
Ky = SslSy(KK)
Tmp = "##" & TmpNm
K = JnComma(Ky)
For J = 0 To UB(Ky)
    Ky(J) = FmtQQ("x.?=a.?", Ky(J), Ky(J))
Next
Jn = Join(Ky, " and ")
A.Execute FmtQQ("Select Distinct ?,Count(*) as Cnt into [?] from [?] group by ? having Count(*)>1", K, Tmp, T, K)
A.Execute FmtQQ("Select x.* into [?] from [?] x inner join [?] a on ?", TarTbl, T, Tmp, Jn)
DbtDrp A, Tmp
End Sub
Sub D(A)
AyDmp A
End Sub
Sub AyDmp(A)
Dim I
If Sz(A) = 0 Then Exit Sub
For Each I In A
    Debug.Print I
Next
End Sub
Function TblFny(A$) As String()
TblFny = DbtFny(CurrentDb, A)
End Function
Function DbtFny_zAutoQuote(A As Database, T$) As String()
Dim O$()
O = DbtFny(A, T)
If DbtIsXls(A, T) Then O = AyQuoteSqBkt(O)
DbtFny_zAutoQuote = O
End Function

Function DbtFny(A As Database, T$) As String()
DbtFny = RsFny(DbtRs(A, T))
End Function
Function DbtIsXls(A As Database, T$) As Boolean
DbtIsXls = HasPfx(A.TableDefs(T).Connect, "Excel")
End Function
Function SplitSpc(A$) As String()
SplitSpc = Split(A, " ")
End Function
Function SqlAny(A$) As Boolean
SqlAny = DbqAny(CurrentDb, A)
End Function
Function RsAny(A As DAO.Recordset) As Boolean
RsAny = Not A.EOF
End Function
Function TblIsExist(T$) As Boolean
TblIsExist = DbHasTbl(CurrentDb, T)
End Function
Sub TblOpn(TblSsl$)
AyDo SslSy(TblSsl), "TblOpn_1"
End Sub
Sub AyDo(A, FunNm$)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run FunNm, I
Next
End Sub
Sub TblOpn_1(T)
DoCmd.OpenTable T
End Sub
Function RplDblSpc$(A)
Dim P%, O$, J%
O = A
While InStr(O, "  ") > 0
    J = J + 1
    If J > 50000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function SslSy(A) As String()
SslSy = SplitSpc(RplDblSpc(Trim(A)))
End Function
Sub ItrNmDo(A, DoFun$)
Dim I
For Each I In A
    Run DoFun, I.Name
Next
End Sub
Sub AcsClsTbl(A As Access.Application)
Dim T
For Each T In A.CodeData.AllTables
    A.DoCmd.Close acTable, T.Name
Next
End Sub

Sub AcsTbl_Cls(A As Access.Application, TT)
'AyNmDoPX A.CodeData.AllTables, "AcsTbl_Cls"
End Sub

Sub ClsTbl()
AcsClsTbl Application
End Sub

Sub TblCls(TT)
AcsTbl_Cls Access.Application, TT
End Sub

Sub TblCls_1(T)
DoCmd.Close acTable, T
End Sub

Sub TblDrp(TT)
DbDrpTbl CurrentDb, TT
End Sub

Sub TblDrp_1(T)
DbtDrp CurrentDb, T
End Sub

Function DbHasQry(A As Database, Q) As Boolean
DbHasQry = DbqAny(A, FmtQQ("Select * from MSysObjects where Name='?' and Type=5", Q))
End Function

Sub DbDrpQry(A As Database, Q)
If DbHasQry(A, Q) Then A.QueryDefs.Delete Q
End Sub

Sub DbCrtQry(A As Database, Q, Sql$)
If Not DbHasQry(A, Q) Then
    Dim QQ As New QueryDef
    QQ.Sql = Sql
    QQ.Name = Q
    A.QueryDefs.Append QQ
Else
    A.QueryDefs(Q).Sql = Sql
End If
End Sub
Function LinShiftTerm$(O$)
Dim A$, P%
A = LTrim(O)
P = InStr(A, " ")
If P = 0 Then
    LinShiftTerm = A
    O = ""
    Exit Function
End If
LinShiftTerm = Left(A, P - 1)
O = LTrim(Mid(A, P + 1))
End Function

Sub LinTTRstAsg(A, OT1$, OT2$, ORst$)
Dim Ay$()
Ay = LinTTRst(A)
OT1 = Ay(0)
OT2 = Ay(1)
ORst = RTrim(Ay(2))
End Sub

Function LinTTRst(A) As String()
Dim O$(2), L$
L = A
O(0) = LinShiftTerm(L)
O(1) = LinShiftTerm(L)
O(2) = L
LinTTRst = O
End Function
Function AyMinus(A, B)
If Sz(B) = 0 Then AyMinus = A: Exit Function
If Sz(A) = 0 Then AyMinus = A: Exit Function
Dim O, I
O = A
Erase O
For Each I In A
    If Not AyHas(B, I) Then Push O, I
Next
AyMinus = O
End Function

Sub DbtRen(A As Database, Fm$, ToTbl$, Optional ReOpnFst As Boolean)
If ReOpnFst Then DbReOpn A
A.TableDefs(Fm).Name = ToTbl
End Sub

Function DbtChkCol(A As Database, T$, LnkColStr$) As String()
Dim Ay() As LnkCol, O$(), Fny$(), J%, Ty As DAO.DataTypeEnum, F$
Ay = LnkColStr_LnkColAy(LnkColStr)
Fny = LnkColAy_ExtNy(Ay)
O = DbtChkFny(A, T, Fny)
If Sz(O) > 0 Then DbtChkCol = O: Exit Function
For J = 0 To UB(Ay)
    F = Ay(J).Extnm
    Ty = Ay(J).Ty
    PushNonEmpty O, DbtChkFldType(A, T, F, Ty)
Next
If Sz(0) > 0 Then
    PushMsgUnderLin O, "Some field has unexpected type"
    DbtChkCol = O
End If
End Function
Function TakAft$(A, S)
Dim P%
P = InStr(A, S)
If P = 0 Then Exit Function
TakAft = Mid(A, P + Len(S))
End Function
Function TakBefOrAll$(A, S)
Dim O$
O = TakBef(A, S)
If O = "" Then
    TakBefOrAll = A
Else
    TakBefOrAll = O
End If
End Function
Function TakAftOrAll$(A, S)
Dim O$
O = TakAft(A, S)
If O = "" Then
    TakAftOrAll = A
Else
    TakAftOrAll = O
End If
End Function


Function TakBef$(A, S)
Dim P%
P = InStr(A, S)
If P = 0 Then Exit Function
TakBef = Left(A, P - 1)
End Function

Function DbtXlsLnkInf(A As Database, T) As XlsLnkInf
Dim Cn$
Cn = DbtCnStr(A, T)
If Not IsPfx(Cn, "Excel") Then Exit Function
With DbtXlsLnkInf
    .IsXlsLnk = True
    .Fx = TakBefOrAll(TakAft(Cn, "DATABASE="), ";")
    .WsNm = A.TableDefs(T).SourceTableName
    If LasChr(.WsNm) <> "$" Then Stop
    .WsNm = RmvLasChr(.WsNm)
End With
End Function

Function AyOfAy_Ay(A)
If Sz(A) = 0 Then AyOfAy_Ay = A: Exit Function
Dim O, J&
O = A(0)
For J = 1 To UB(A)
    PushAy O, A(J)
Next
AyOfAy_Ay = O
End Function
Function ISpecINm$(A$)
ISpecINm = LinT1(A)
End Function
Sub LSpecDmp(A)
Debug.Print RplVBar(A)
End Sub
Function LSpecLy(A) As String()
Const L2Spec$ = ">GLAnp |" & _
    "Whs    Txt Plant |" & _
    "Loc    Txt [Storage Location]|" & _
    "Sku    Txt Material |" & _
    "PstDte Txt [Posting Date] |" & _
    "MovTy  Txt [Movement Type]|" & _
    "Qty    Txt Quantity|" & _
    "BchNo  Txt Batch |" & _
    "Where Plant='8601' and [Storage Location]='0002' and [Movement Type] like '6*'"
End Function
Function HasPfx(A, Pfx$) As Boolean
HasPfx = Left(A, Len(Pfx)) = Pfx
End Function
Sub LSpecAsg(A, Optional OTblNm$, Optional OLnkColStr$, Optional OWhBExpr$)
Dim Ay$()
Ay = AyTrim(SplitVBar(A))
OTblNm = AyShift(Ay)
If LinT1(AyLasEle(Ay)) = "Where" Then
    OWhBExpr = LinRmvTerm(Pop(Ay))
Else
    OWhBExpr = ""
End If
OLnkColStr = JnVBar(Ay)
End Sub
Function Pop(A)
Pop = AyLasEle(A)
AyRmvLasEle A
End Function
Sub AyRmvLasEle(A)
If Sz(A) = 1 Then
    Erase A
    Exit Sub
End If
ReDim Preserve A(UB(A) - 1)
End Sub
Function JnVBar$(A)
JnVBar = Join(A, "|")
End Function
Sub LSpecAy_Asg(A$(), OTny$(), OLnkColStrAy$(), OWhBExprAy$())
Dim U%, J%
U = UB(A)
ReDim OTny(U)
ReDim OLnkColStrAy(U)
ReDim OWhBExprAy(U)
For J = 0 To U
    LSpecAsg A(J), OTny(J), OLnkColStrAy(J), OWhBExprAy(J)
Next
End Sub

Function DbImp(A As Database, LSpec$()) As String()
Dim O$(), J%, T$(), L$(), W$(), U%
LSpecAy_Asg LSpec, T, L, W
U = UB(LSpec)
For J = 0 To U
    PushAy O, DbtChkCol(A, T(J), L(J))
Next
If Sz(O) > 0 Then DbImp = O: Exit Function
For J = 0 To U
    DbtImpMap A, T(J), L(J), W(J)
Next
DbImp = O
End Function

Function DbtMissFny_Er(A As Database, T$, MissFny$(), ExistingFny$()) As String()
Dim X As XlsLnkInf, O$(), I
If Sz(MissFny) = 0 Then Exit Function
X = DbtXlsLnkInf(A, T)
If X.IsXlsLnk Then
    Push O, "Excel File       : " & X.Fx
    Push O, "Worksheet        : " & X.WsNm
    PushUnderLin O
    For Each I In ExistingFny
        Push O, "Worksheet Column : " & QuoteSqBkt(CStr(I))
    Next
    PushUnderLin O
    For Each I In MissFny
        Push O, "Missing Column   : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Columns are missing"
Else
    Push O, "Database : " & A.Name
    Push O, "Table    : " & T
    For Each I In MissFny
        Push O, "Field    : " & QuoteSqBkt(CStr(I))
    Next
    PushMsgUnderLinDbl O, "Above Fields are missing"
End If
DbtMissFny_Er = O
End Function

Function DbtChkFny(A As Database, T$, ExpFny$()) As String()
Dim Miss$(), TFny$(), O$(), I
TFny = DbtFny(A, T)
Miss = AyMinus(ExpFny, TFny)
DbtChkFny = DbtMissFny_Er(A, T, Miss, TFny)
End Function
Function QuoteSqBkt$(A$)
QuoteSqBkt = "[" & A & "]"
End Function
Function PushMsgUnderLin(O$(), M$)
Push O, M
Push O, UnderLin(M)
End Function
Function PushUnderLin(O$())
Push O, UnderLin(AyLasEle(O))
End Function
Function PushUnderLinDbl(O$())
Push O, UnderLinDbl(AyLasEle(O))
End Function
Function PushMsgUnderLinDbl(O$(), M$)
Push O, M
Push O, UnderLinDbl(M)
End Function
Function DaoTy_ShtTy$(A As DAO.DataTypeEnum)
Dim O$
Select Case A
Case DAO.DataTypeEnum.dbByte: O = "Byt"
Case DAO.DataTypeEnum.dbLong: O = "Lng"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbDate: O = "Dte"
Case DAO.DataTypeEnum.dbText: O = "Txt"
Case DAO.DataTypeEnum.dbBoolean: O = "Yes"
Case DAO.DataTypeEnum.dbDouble: O = "Dbl"
Case Else: Stop
End Select
DaoTy_ShtTy = O
End Function
Function DaoShtTy_Ty(A$) As DAO.DataTypeEnum
Dim O As DAO.DataTypeEnum
Select Case A
Case "Byt": O = DAO.DataTypeEnum.dbByte
Case "Lng": O = DAO.DataTypeEnum.dbLong
Case "Int": O = DAO.DataTypeEnum.dbInteger
Case "Dte": O = DAO.DataTypeEnum.dbDate
Case "Txt": O = DAO.DataTypeEnum.dbText
Case "Yes": O = DAO.DataTypeEnum.dbBoolean
Case "Dbl": O = DAO.DataTypeEnum.dbDouble
Case Else: Stop
End Select
DaoShtTy_Ty = O
End Function
Function DftFfnAy(FfnAy0) As String()
Select Case True
Case IsStr(FfnAy0): DftFfnAy = ApSy(FfnAy0)
Case IsSy(FfnAy0): DftFfnAy = FfnAy0
Case IsArray(FfnAy0): DftFfnAy = AySy(FfnAy0)
End Select
End Function
Property Get FfnCpyToPthIfDif(FfnAy0, Pth$) As String()
Const M_Sam$ = "File is same the one in Path."
Const M_Copied$ = "File is copied to Path."
Const M_NotFnd$ = "File not found, cannot copy to Path."
PthEns Pth
Dim B$, Ay$(), I, O$(), M$(), Msg$
Ay = DftFfnAy(FfnAy0): If Sz(Ay) = 0 Then Exit Property
For Each I In Ay
    Select Case True
    Case FfnIsExist(CStr(I))
        B = Pth & FfnFn(I)
        Select Case True
        Case FfnIsSam(B, CStr(I))
            Msg = M_Sam: GoSub Prt
        Case Else
            Fso.CopyFile I, B, True
            Msg = M_Copied: GoSub Prt
        End Select
    Case Else
        Msg = M_NotFnd: GoSub Prt
        Push O, "File : " & I
    End Select
Next
If Sz(O) > 0 Then
    PushMsgUnderLinDbl O, "Above files not found"
    FfnCpyToPthIfDif = O
End If
Exit Property
Prt:
    Debug.Print FmtQQ("FfnCpyToPthIfDif: ? Path=[?] File=[?]", Msg, Pth, I)
    Return
End Property
Function FfnIsSamMsg(A$, B$, Sz&, Tim$, Optional Msg$) As String()
Dim O$()
Push O, "File 1   : " & A
Push O, "File 2   : " & B
Push O, "File Size: " & Sz
Push O, "File Time: " & Tim
Push O, "File 1 and 2 have same size and time"
If Msg <> "" Then Push O, Msg
FfnIsSamMsg = O
End Function
Function FfnIsSam(A$, B$) As Boolean
If FfnTim(A) <> FfnTim(B) Then Exit Function
If FfnSz(A) <> FfnSz(B) Then Exit Function
FfnIsSam = True
End Function
Function FfnSz&(A$)
If FfnIsExist(A) Then
    FfnSz = FileLen(A)
Else
    FfnSz = -1
End If
End Function
Function FfnTim(A$) As Date
If FfnIsExist(A) Then FfnTim = FileDateTime(A)
End Function
Function AyTrim(A) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Trim(A(J))
Next
AyTrim = O
End Function
Function DbtChkFldType$(A As Database, T$, F, Ty As DAO.DataTypeEnum)
Dim ActTy As DAO.DataTypeEnum
ActTy = A.TableDefs(T).Fields(F).Type
If ActTy <> Ty Then
    DbtChkFldType = FmtQQ("Table[?] field[?] should have type[?], but now it has type[?]", T, F, DaoTy_ShtTy(Ty), DaoTy_ShtTy(ActTy))
End If
End Function
Function OyPrpSy(A, PrpNm$) As String()
If Sz(A) = 0 Then Exit Function
OyPrpSy = ItrPrpSy(A, PrpNm)
End Function
Function OyPrpInto(A, PrpNm$, OInto)
If Sz(A) = 0 Then Exit Function
OyPrpInto = ItrPrpInto(A, PrpNm, OInto)
End Function
Function LnkColAy_ExtNy(A() As LnkCol) As String()
LnkColAy_ExtNy = OyPrpSy(A, "Extnm")
End Function
Function LnkColAy_Ny(A() As LnkCol) As String()
LnkColAy_Ny = OyPrpSy(A, "Nm")
End Function
Sub WbVdtOupNy(A As Workbook, OupNy$())
Dim O$(), N$, B$(), WsCdNy$()
WsCdNy = WbWsCdNy(A)
O = AyMinus(AyAddPfx(OupNy, "WsO"), WsCdNy)
If Sz(O) > 0 Then
    N = "OupNy":  B = OupNy:  GoSub Dmp
    N = "WbCdNy": B = WsCdNy: GoSub Dmp
    N = "Mssing": B = O:      GoSub Dmp
    Stop
    Exit Sub
End If
Exit Sub
Dmp:
Debug.Print UnderLin(N)
Debug.Print N
Debug.Print UnderLin(N)
AyDmp B
Return
End Sub
Function RsDrs(A As DAO.Recordset) As Drs
Dim Fny$(), Dry()
Fny = RsFny(A)
Dry = RsDry(A)
Set RsDrs = Drs(Fny, Dry)
End Function
Function RsDr(A As DAO.Recordset) As Variant()
RsDr = FldsDr(A.Fields)
End Function
Function RsDry(A As DAO.Recordset) As Variant()
Dim O()
Push O, RsFny(A)
With A
    While Not .EOF
        Push O, RsDr(A)
        .MoveNext
    Wend
End With
RsDry = O
End Function
Function LoHasFny(A As ListObject, Fny$()) As Boolean
Dim Miss$(), FnyzLo$()
FnyzLo = LoFny(A)
Miss = AyMinus(Fny, FnyzLo)
If Sz(Miss) > 0 Then Exit Function
LoHasFny = True
End Function
Function WsFstLo(A As Worksheet) As ListObject
Set WsFstLo = ItrFstItm(A.ListObjects)
End Function
Function ItrFstItm(A)
Dim I
For Each I In A
    Asg I, ItrFstItm
Next
End Function
Function DrsNRow&(A As Drs)
DrsNRow = Sz(A.Dry)
End Function
Function SqAddSngQuote(A)
Dim NC%, C%, R&, O
O = A
NC = UBound(A, 2)
For R = 1 To UBound(A, 1)
    For C = 1 To NC
        If IsStr(O(R, C)) Then
            O(R, C) = "'" & O(R, C)
        End If
    Next
Next
SqAddSngQuote = O
End Function
Sub FldsPutSq(A As DAO.Fields, Sq, R&)
Dim C%, F As DAO.Field
C = 1
For Each F In A
    Sq(R, C) = F.Value
    C = C + 1
Next
End Sub
Function RsSq(A As DAO.Recordset) As Variant()
RsSq = DrySq(RsDry(A))
End Function
Sub DbtPutLo(A As Database, T$, Lo As ListObject)
Dim Sq(), Drs As Drs, Rs As DAO.Recordset
Set Rs = DbtRs(A, T)
If Not AyIsEq(RsFny(Rs), LoFny(Lo)) Then
    Debug.Print "--"
    Debug.Print "Rs"
    Debug.Print "--"
    AyDmp RsFny(Rs)
    Debug.Print "--"
    Debug.Print "Lo"
    Debug.Print "--"
    AyDmp LoFny(Lo)
    Stop
End If
Sq = SqAddSngQuote(RsSq(Rs))
LoMin Lo
SqPutAt Sq, Lo.DataBodyRange
End Sub
Sub LoEnsNRow(A As ListObject, NRow&)
LoMin A
Exit Sub
If NRow > 1 Then
    Debug.Print A.InsertRowRange.Address
    Stop
End If
End Sub
Function DrsCol(A As Drs, F) As Variant()
DrsCol = DrsColInto(A, F, EmpAy)
End Function
Function AyIx&(A, M)
Dim J&
For J = 0 To UB(A)
    If A(J) = M Then AyIx = J: Exit Function
Next
AyIx = -1
End Function
Function LoSy(A As ListObject, ColNm$) As String()
Dim Sq()
Sq = A.ListColumns(ColNm).DataBodyRange.Value
LoSy = SqColSy(Sq, 1)
End Function
Function LoFny(A As ListObject) As String()
LoFny = ItrNy(A.ListColumns)
End Function
Sub AyPutLoCol(A, Lo As ListObject, ColNm$)
Dim At As Range, C As ListColumn, R As Range
'AyDmp LoFny(Lo)
'Stop
Set C = Lo.ListColumns(ColNm)
Set R = C.DataBodyRange
Set At = R.Cells(1, 1)
AyPutCol A, At
End Sub
Function AySqH(A) As Variant()
Dim O(), N&, J&
N = Sz(A)
If N = 0 Then Exit Function
ReDim Sq(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = A(J - 1)
Next
AySqH = O
End Function
Function AySqV(A) As Variant()
Dim O(), N&, J&
N = Sz(A)
If N = 0 Then Exit Function
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = A(J - 1)
Next
AySqV = O
End Function
Sub AyPutCol(A, At As Range)
Dim Sq()
Sq = AySqV(A)
RgReSz(At, Sq).Value = Sq
End Sub
Sub AyPutRow(A, At As Range)
Dim Sq()
Sq = AySqH(A)
RgReSz(At, Sq).Value = Sq
End Sub
Function DrsColInto(A As Drs, F, OInto)
Dim O, Ix%, Dry(), Dr
Ix = AyIx(A.Fny, F): If Ix = -1 Then Stop
O = OInto
Erase O
Dry = A.Dry
If Sz(Dry) = 0 Then DrsColInto = O: Exit Function
For Each Dr In Dry
    Push O, Dr(Ix)
Next
DrsColInto = O
End Function

Function DrsColSy(A As Drs, F) As String()
DrsColSy = DrsColInto(A, F, EmpSy)
End Function

Sub DbtDrp(A As Database, TT)
Dim Tny$(), T
Tny = CvNy(TT)
For Each T In Tny
    If DbHasTbl(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
Next
End Sub

Function DbtLnk(A As Database, T$, S$, Cn$) As String()
On Error GoTo X
Dim TT As New DAO.TableDef
DbDrpTbl A, T
With TT
    .Connect = Cn
    .Name = T
    .SourceTableName = S
    A.TableDefs.Append TT
End With
Exit Function
X:
Dim Er$
Er = Err.Description
Debug.Print Er
Dim O$(), M$
M = "Cannot create Table in Database from Source by Cn with Er from system"
Push O, "Program  : DbtLnk"
Push O, "Database : " & A.Name
Push O, "Table    : " & T
Push O, "Source   : " & S
Push O, "Cn       : " & Cn
Push O, "Er       : " & Er
PushMsgUnderLin O, M
DbtLnk = O
End Function
Function TblLnk(T$, S$, Cn$) As String()
TblLnk = DbtLnk(CurrentDb, T, S, Cn)
End Function
Function WbWsCd(A As Workbook, WsCdNm$) As Worksheet
Set WbWsCd = ItrFstPrpEq(A.Sheets, "CodeName", WsCdNm)
End Function
Function WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Function
Function WbWs(A As Workbook, WsNm) As Worksheet
Set WbWs = A.Sheets(WsNm)
End Function
Function FxWb(A) As Workbook
Set FxWb = NewXls.Workbooks.Open(A)
End Function
Function WsLo(A As Worksheet, LoNm$) As ListObject
Set WsLo = A.ListObjects(LoNm)
End Function
Function TblRg(A$, At As Range) As Range
Set TblRg = DbtRg(CurrentDb, A, At)
End Function
Function DbtRg(A As Database, T$, At As Range) As Range
Set DbtRg = SqRg(DbtSq(A, T), At)
End Function
Function AyAddAp(ParamArray Ap())
Dim Av(), O, J%
O = Ap(0)
Av = Ap
For J = 1 To UB(Av)
    PushAy O, Av(J)
Next
AyAddAp = O
End Function
Function AlignL$(A, W%)
AlignL = A & Space(W - Len(A))
End Function

Function AyMapXPSy(A, MapXPFunNm$, P) As String()
AyMapXPSy = AyMapXPInto(A, MapXPFunNm, P, EmpSy)
End Function

Function AyMapXPInto(A, MapXPFunNm$, P, OInto)
Dim O, J&
O = OInto
Erase O
If Sz(A) = 0 Then AyMapXPInto = O: Exit Function
ReDim O(UB(A))
For J = 0 To UB(A)
    Asg Run(MapXPFunNm, A(J), P), O(J)
Next
AyMapXPInto = O
End Function

Function AyAlignL(A) As String()
AyAlignL = AyMapXPSy(A, "AlignL", AyWdt(A))
End Function
Function LSpecLnkColStr$(A)
Dim L$
LSpecAsg A, , L
LSpecLnkColStr = L
End Function
Function LnkColAy_ImpSql$(A() As LnkCol, T$, Optional WhBExpr$)
If FstChr(T) <> ">" Then
    Debug.Print "T must have first char = '>'"
    Stop
End If
Dim Ny$(), ExtNy$(), J%, O$(), S$, N$(), E$()
Ny = LnkColAy_Ny(A)
ExtNy = LnkColAy_ExtNy(A)
N = AyAlignL(Ny)
E = AyAlignL(AyQuoteSqBkt(ExtNy))
Erase O
For J = 0 To UB(Ny)
    If ExtNy(J) = Ny(J) Then
        Push O, FmtQQ("     ?    ?", Space(Len(E(J))), N(J))
    Else
        Push O, FmtQQ("     ? As ?", E(J), N(J))
    End If
Next
S = Join(O, "," & vbCrLf)
LnkColAy_ImpSql = FmtQQ("Select |?| Into [#I?]| From [?] |?", S, RmvFstChr(T), T, PWh(WhBExpr))
End Function
Sub WbMinLo(A As Workbook)
ItrDo A.Sheets, "WsMinLo"
End Sub
Sub WsMinLo(A As Worksheet)
If A.CodeName = "WsIdx" Then Exit Sub
ItrDo A.ListObjects, "LoMin"
End Sub
Sub LoMin(A As ListObject)
Dim R1 As Range, R2 As Range
Set R1 = A.DataBodyRange
If R1.Rows.Count >= 2 Then
    Set R2 = RgRR(R1, 2, R1.Rows.Count)
    R2.Delete
End If
End Sub
Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgCRR(A, 1, R1, R2).EntireRow
End Function
Sub FxMinLo(A)
Dim Wb As Workbook
Set Wb = FxWb(A)
WbMinLo Wb
Wb.Save
Wb.Close
End Sub
Sub PcRfh(A As PivotCache)
A.MissingItemsLimit = xlMissingItemsNone
A.Refresh
End Sub
Sub ItrDo(A, DoNm$)
Dim I
For Each I In A
    Run DoNm, I
Next
End Sub
Sub ItrDoXP(A, DoXPNm$, P)
Dim I
For Each I In A
    Run DoXPNm, I, P
Next
End Sub
Function IsDev() As Boolean
Static X As Boolean, Y As Boolean
If Not X Then
    X = True
    Y = Not PthIsExist("N:\SAPAccessReports\")
End If
IsDev = Y
End Function

Sub FunPAy_Do(A, P)
Dim FunP
For Each FunP In A
    Run CStr(FunP), P
Next
End Sub
Function OupFx_Crt$(A$)
OupFx_Crt = AttExp("Tp", A)
End Function
Sub OupFx_Gen(OupFx$, Fb$, ParamArray WbFmtrAp())
Dim Av()
Av = WbFmtrAp
TpWrtFfn OupFx
WbFmt FxRfh(OupFx, Fb), Av
End Sub
Function FxRfh(A, Fb$) As Workbook
Set FxRfh = WbRfh(FxWb(A), Fb)
End Function
Sub WbFmt(A As Workbook, WbFmtrAv())
If True Then
    FunPAy_Do WbFmtrAv, A
Else
    Dim J%
    For J = 0 To UB(WbFmtrAv)
        Run WbFmtrAv(J), A
    Next
End If
WbMax(WbVis(A)).Save
End Sub
Sub TpGenFx(TpFx$, OupFx$, Fb$, ParamArray WbFmtrAp())
Dim Av()
Av = WbFmtrAp
FfnCpy TpFx, OupFx
WbFmt FxRfh(OupFx, Fb), Av
End Sub

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function
Function CvLo(A) As ListObject
Set CvLo = A
End Function
Function DbOupTny(A As Database) As String()
DbOupTny = DbqSy(A, "Select Name from MSysObjects where Name like '@*' and Type =1")
End Function
Function ObjNm$(O)
ObjNm = CallByName(O, "Name", VbGet)
End Function

Function ObjHasNmPfx(O, NmPfx$) As Boolean
ObjHasNmPfx = HasPfx(ObjNm(O), NmPfx)
End Function

Function OyWhNmHasPfx(A, Pfx$)
OyWhNmHasPfx = OyWhPredXP(A, "ObjHasNmPfx", Pfx)
End Function

Function OyWhPredXP(A, XP$, P)
Dim O, X
O = A
Erase O
For Each X In A
    If Run(XP, X, P) Then
        PushObj A, X
    End If
Next
OyWhPredXP = O
End Function

Function WbOupLoAy(A As Workbook) As ListObject()
WbOupLoAy = OyWhNmHasPfx(WbLoAy(A), "T_")
End Function

Sub FbRplWbLo(Fb$, A As Workbook)
Dim I, Lo As ListObject, Db As Database
Set Db = FbDb(Fb)
For Each I In WbOupLoAy(A)
    Set Lo = I
    DbtRplLo Db, "@" & Mid(Lo.Name, 3), Lo
Next
Db.Close
Set Db = Nothing
End Sub

Function WbRfh(A As Workbook, Optional Fb$) As Workbook
ItrDoXP A.Connections, "WcRfh", Fb
ItrDo A.PivotCaches, "PcRfh"
ItrDo A.Sheets, "WsRfh"
Set WbRfh = A
End Function
Sub WbDltWc(A As Workbook)
ItrDo A.Connections, "WcDlt"
End Sub
Sub ZZ_RplBet()
Dim A$, Exp$, By$, S1$, S2$
S1 = "Data Source="
S2 = ";"
A = "aa;Data Source=???;klsdf"
By = "xx"
Exp = "aa;Data Source=xx;klsdf"
GoSub Tst
Exit Sub
Tst:
Dim Act$
Act = RplBet(A, By, S1, S2)
Debug.Assert Exp = Act
Return
End Sub
Function RplBet$(A$, By$, S1$, S2$)
Dim P1%, P2%, B$, C$

P1 = InStr(A, S1)
If P1 = 0 Then Stop
P2 = InStr(P1 + Len(S1), CStr(A), S2)
If P2 = 0 Then Stop
B = Left(A, P1 + Len(S1) - 1)
C = Mid(A, P2 + Len(S2) - 1)
RplBet = B & By & C
End Function
Sub A()
ZZ_FbWb_zExpOupTbl
End Sub
Function FbWcStr$(A)
FbWcStr = FbOleCnStr(A)
'FbWcStr = FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False", A)
End Function
Sub WcRfhCnStr(A As WorkbookConnection, Optional Fb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
If Fb = "" Then Exit Sub
Dim Cn$
Const Ver$ = "0.0.1"
Select Case Ver
Case "0.0.1"
    Dim S$
    S = A.OLEDBConnection.Connection
    Cn = RplBet(S, CStr(Fb), "Data Source=", ";")
Case "0.0.2"
    Cn = FbWcStr(Fb)
End Select
A.OLEDBConnection.Connection = Cn
End Sub
Sub WcRfh(A As WorkbookConnection, Optional Fb$)
If IsNothing(A.OLEDBConnection) Then Exit Sub
WcRfhCnStr A, Fb
A.OLEDBConnection.BackgroundQuery = False
A.OLEDBConnection.Refresh
End Sub
Sub WcDlt(A As WorkbookConnection)
A.Delete
End Sub

Sub WsRfh(A As Worksheet)
ItrDo A.QueryTables, "QtRfh"
ItrDo A.PivotTables, "PtRfh"
End Sub

Sub QtRfh(A As Excel.QueryTable)
A.BackgroundQuery = False
A.Refresh
End Sub
Sub PtRfh(A As Excel.PivotTable)
A.Update
End Sub
Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
Function LoVis(A As ListObject) As ListObject
XlsVis A.Application
Set LoVis = A
End Function
Function WsVis(A As Worksheet)
XlsVis A.Application
Set WsVis = A
End Function
Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub
Function SqPutAt(A, At As Range) As Range
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqPutAt = O
End Function
Function RgWs(A As Range) As Worksheet
Set RgWs = A.Parent
End Function
Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function
Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function
Function RgReSz(A As Range, Sq) As Range
Set RgReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function
Sub ZZ_TblSq()
Dim A()
A = TblSq("@Oup")
Stop
End Sub
Function NewWb(Optional WsNm$ = "Sheet1") As Workbook
Dim O As Workbook, Ws As Worksheet
Set O = NewXls.Workbooks.Add
Set Ws = WbFstWs(O)
If Ws.Name <> WsNm Then Ws.Name = WsNm
Set NewWb = O
End Function
Function WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Function
Function NewWs(Optional WsNm$ = "Sheet") As Worksheet
Set NewWs = WbFstWs(NewWb(WsNm))
End Function
Function NewA1(Optional WsNm$ = "Sheet1") As Range
Set NewA1 = WsA1(NewWs(WsNm))
End Function
Function SqNewA1(A, Optional WsNm$ = "Data") As Range
Dim A1 As Range
Set A1 = NewA1(WsNm)
Set SqNewA1 = SqPutAt(A, A1)
End Function
Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function
Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function
Function RgA1LasCell(A As Range) As Range
Dim L As Range, R, C
Set L = A.SpecialCells(xlCellTypeLastCell)
R = L.Row
C = L.Column
Set RgA1LasCell = WsRCRC(RgWs(A), A.Row, A.Column, R, C)
End Function
Function RgLo(A As Range, Optional LoNm$) As ListObject
Dim O As ListObject
Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , XlYesNoGuess.xlYes)
'LoAutoFit O
If LoNm <> "" Then O.Name = LoNm
Set RgLo = O
End Function
Function RgVis(A As Range) As Range
XlsVis A.Application
Set RgVis = A
End Function
Sub DbtWrtFx(A As Database, TT, Fx$)
DbttWb(A, TT).SaveAs Fx
End Sub
Sub WsClrLo(A As Worksheet)
Dim Ay() As ListObject, J%
Ay = ItrAy(A.ListObjects, Ay)
For J = 0 To UB(Ay)
    Ay(J).Delete
Next
End Sub
Sub TblWrtFx(TT, Fx$)
DbtWrtFx CurrentDb, TT, Fx
End Sub
Function WbAddWs(A As Workbook, Optional WsNm, Optional BefWsNm$, Optional AftWsNm$) As Worksheet
Dim O As Worksheet, Bef As Worksheet, Aft As Worksheet
WbDltWs A, WsNm
Select Case True
Case BefWsNm <> ""
    Set Bef = A.Sheets(BefWsNm)
    Set O = A.Sheets.Add(Bef)
Case AftWsNm <> ""
    Set Aft = A.Sheets(AftWsNm)
    Set O = A.Sheets.Add(, Aft)
Case Else
    Set O = A.Sheets.Add
End Select
O.Name = WsNm
Set WbAddWs = O
End Function
Sub WbDltWs(A As Workbook, WsNm)
If WbHasWs(A, WsNm) Then
    A.Application.DisplayAlerts = False
    WbWs(A, WsNm).Delete
    A.Application.DisplayAlerts = True
End If
End Sub
Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function

Function WbHasWs(A As Workbook, WsNm) As Boolean
WbHasWs = ItrHasNm(A.Sheets, WsNm)
End Function

Sub FfnCpy(A, ToFfn$, Optional OvrWrt As Boolean)
If OvrWrt Then FfnDlt ToFfn
FileSystem.FileCopy A, ToFfn
End Sub

Sub FfnDlt(A)
If FfnIsExist(A) Then Kill A
End Sub

Function PthIsExist(A) As Boolean
On Error Resume Next
PthIsExist = Dir(A, vbDirectory) <> ""
End Function
Function FfnIsExist(A) As Boolean
On Error Resume Next
FfnIsExist = Dir(A) <> ""
End Function
Function TTWb(TT, Optional UseWc As Boolean) As Workbook
Set TTWb = DbttWb(CurrentDb, TT, UseWc)
End Function
Function DbttWb(A As Database, TT, Optional UseWc As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
Set DbttWb = WbAddDbtt(O, A, TT, UseWc)
WbWs(O, "Sheet1").Delete
End Function
Function WbA1(A As Workbook, Optional WsNm) As Range
Set WbA1 = WsA1(WbAddWs(A, WsNm))
End Function
Function DbtAt_Lo(A As Database, T$, At As Range, Optional UseWc As Boolean) As ListObject
Dim N$, Q As QueryTable
N = TnLoNm(T)
If UseWc Then
    Set Q = RgWs(At).ListObjects.Add(SourceType:=0, Source:=FbAdoCnStr(A.Name), Destination:=At).QueryTable
    With Q
        .CommandType = xlCmdTable
        .CommandText = T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = T
        .Refresh BackgroundQuery:=False
    End With
    Exit Function
End If
Set DbtAt_Lo = RgLo(DbtRg(A, T, At), N)
End Function
Function LoWb(A As ListObject) As Workbook
Set LoWb = LoWs(A).Parent
End Function
Function WbAddDbt(A As Workbook, Db As Database, T$, Optional UseWc As Boolean) As Workbook
Set WbAddDbt = LoWb(DbtAt_Lo(Db, T, WbA1(A, T), UseWc))
End Function
Function TnLoNm$(TblNm)
TnLoNm = "T_" & RmvFstNonLetter(TblNm)
End Function
Sub AyDoPPXP(A, PPXP$, P1, P2, P3)
Dim X
For Each X In A
    Run PPXP, P1, P2, X, P3
Next
End Sub

Function WbAddDbtt(A As Workbook, Db As Database, TT, Optional UseWc As Boolean) As Workbook
AyDoPPXP CvTT(TT), "WbAddDbt", A, Db, UseWc
Set WbAddDbtt = A
End Function

Function DbqSy(A As Database, Sql) As String()
DbqSy = RsSy(A.OpenRecordset(Sql))
End Function
Function DbStru$(A As Database)
DbStru = DbttStru(A, DbTny(A))
End Function

Property Get Tny() As String()
Tny = DbTny(CurrentDb)
End Property
Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Name not Like 'MSys*' and Name not Like 'f_????????????????????????????????_*'")
End Function
Function IsPfx(A$, Pfx$) As Boolean
IsPfx = Left(A, Len(Pfx)) = Pfx
End Function
Function DbtNRec&(A As Database, T)
DbtNRec = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function
Function DbtCsv(A As Database, T) As String()
DbtCsv = RsCsvLy(DbtRs(A, T))
End Function
Function TblNm_LoNm$(A$)
TblNm_LoNm = "T_" & RmvFstNonLetter(A)
End Function
Function DbtLo(A As Database, T$, At As Range) As ListObject
Set DbtLo = SqAt_Lo(DbtSq(A, T), At, TblNm_LoNm(T))
End Function
Function DSpecNm$(A)
DSpecNm = TakAftDotOrAll(LinT1(A))
End Function
Function TakAftDotOrAll$(A)
TakAftDotOrAll = TakAftOrAll(A, ".")
End Function
Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function
Function DbtRs(A As Database, T) As DAO.Recordset
Set DbtRs = A.OpenRecordset(T)
End Function
Function TblRs(T) As DAO.Recordset
Set TblRs = DbtRs(CurrentDb, T)
End Function
Sub TimFn(FnNm$)
Dim A!, B!
A = Timer
Run FnNm
B = Timer
Debug.Print FnNm, B - A
End Sub
Function RsCsvLyByFny0(A As DAO.Recordset, Fny0) As String()
Dim Fny$(), Flds As Fields, F
Dim O$(), J&, I%, UFld%, Dr()
Fny = CvNy(Fny0)
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    Set Flds = A.Fields
    For Each F In Fny
        Dr(I) = VarCsv(Flds(F).Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLyByFny0 = O
End Function
Function RsCsvLy(A As DAO.Recordset) As String()
Dim O$(), J&, I%, UFld%, Dr(), F As DAO.Field
UFld = A.Fields.Count - 1
While Not A.EOF
    J = J + 1
    If J Mod 5000 = 0 Then Debug.Print "RsCsvLy: " & J
    If J > 100000 Then Stop
    ReDim Dr(UFld)
    I = 0
    For Each F In A.Fields
        Dr(I) = VarCsv(F.Value)
        I = I + 1
    Next
    Push O, Join(Dr, ",")
    A.MoveNext
Wend
RsCsvLy = O
End Function

Function TblNRow&(T$, Optional WhBExpr$)
TblNRow = DbtNRow(CurrentDb, T, WhBExpr)
End Function
Function PWh$(WhBExpr$)
If WhBExpr = "" Then Exit Function
PWh = PSep & "Where" & PSep1 & WhBExpr
End Function
Function DbtNRow&(A As Database, T$, Optional WhBExpr$)
Dim S$
S = "Select Count(*)" & PFm(T) & PWh(WhBExpr)
DbtNRow = DbqLng(A, S)
End Function
Function TblNCol&(T)
TblNCol = DbtNCol(CurrentDb, T)
End Function
Function DbtNCol&(A As Database, T)
DbtNCol = A.OpenRecordset(T).Fields.Count
End Function
Function TblSq(A$) As Variant()
TblSq = DbtSq(CurrentDb, A)
End Function
Function DbtSq(A As Database, T$, Optional ReSeqSpec$) As Variant()
Dim Q$
Q = QSel(T, ReSeqSpec_Fny(ReSeqSpec))
DbtSq = RsSq(DbqRs(A, Q))
End Function
Sub ZZ_QSel()
Debug.Print QSel("A")
End Sub
Function QSel$(T, Optional Fny0, Optional FldExprDic As Dictionary)
QSel = PSel(Fny0, FldExprDic) & PFm(T)
End Function
Function PFm$(T)
PFm = PSep & "From [" & T & "]"
End Function
Function PFmAlias$(T$, Alias$)
PFmAlias = PFm(T) & " " & Alias
End Function
Function PSel$(Fny0, Optional FldExprDic As Dictionary)
Dim Fny$()
Fny = CvNy(Fny0)
If Sz(Fny) = 0 Then
    PSel = "Select *"
    Exit Function
End If
PSel = "Select " & JnComma(CvNy(Fny0))
End Function
Function PAddCol$(Fny0, FldDfnDic As Dictionary)
Dim Fny$(), O$(), J%
Fny = CvNy(Fny0)
ReDim O(UB(Fny))
For J = 0 To UB(Fny)
    O(J) = Fny(J) & " " & FldDfnDic(Fny(J))
Next
PAddCol = PSep & "Add Column " & JnComma(O)
End Function
Function FxWs(A, Optional WsNm$ = "Data") As Worksheet
Set FxWs = WbWs(FxWb(A), WsNm)
End Function
Sub FldsPutSq1(A As DAO.Fields, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
DrPutSq FldsDr(A), Sq, R, NoTxtSngQ
End Sub
Sub DrPutSq(A, Sq, Optional R& = 1, Optional NoTxtSngQ As Boolean)
Dim J%, I
If NoTxtSngQ Then
    For Each I In A
        J = J + 1
        Sq(R, J) = I
    Next
    Exit Sub
End If
For Each I In A
    J = J + 1
    If IsStr(I) Then
        Sq(R, J) = "'" & I
    Else
        Sq(R, J) = I
    End If
Next
End Sub
Sub RsPutSq(A As DAO.Recordset, Sq, R&, Optional NoTxtSngQ As Boolean)
FldsPutSq1 A.Fields, Sq, R, NoTxtSngQ
End Sub
Function WsRCC(A As Worksheet, R, C1, C2) As Range
Set WsRCC = WsRCRC(A, R, C1, R, C2)
End Function
Function WsCC(A As Worksheet, C1, C2) As Range
Set WsCC = WsRCC(A, 1, C1, C2).EntireColumn
End Function
Function WsRR(A As Worksheet, R1&, R2&) As Range
Set WsRR = A.Rows(R1 & ":" & R2)
End Function
Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Cells(1, 1)
End Function
Function FxLo(A$, Optional WsNm$ = "Data", Optional LoNm$ = "Data") As ListObject
Set FxLo = WsLo(WbWs(FxWb(A), WsNm), LoNm)
End Function
Function TblCnStr$(T)
TblCnStr = CurrentDb.TableDefs(T).Connect
End Function
Function DbqLng&(A As Database, Sql)
DbqLng = DbqV(A, Sql)
End Function
Function SqlLng&(A)
SqlLng = DbqLng(CurrentDb, A)
End Function
Function SqlV(A)
SqlV = DbqV(CurrentDb, A)
End Function
Function DbqV(A As Database, Sql)
DbqV = A.OpenRecordset(Sql).Fields(0).Value
End Function
Function TblNRec&(A)
TblNRec = SqlLng(FmtQQ("Select Count(*) from [?]", A))
End Function
Function ErzFileNotFound(FfnAy0) As String()
Dim Ay$(), I, O$()
Ay = DftFfnAy(FfnAy0)
If Sz(Ay) = 0 Then Exit Function
For Each I In Ay
    If Not FfnIsExist(CStr(I)) Then
        Push O, "File: " & I
        PushMsgUnderLin O, "Above file not found"
    End If
Next
ErzFileNotFound = O
End Function
Function DbtLnkFx(A As Database, T$, Fx$, Optional WsNm$ = "Sheet1") As String()
Dim O$()
O = ErzFileNotFound(Fx)
If Sz(O) > 0 Then
    DbtLnkFx = O
    Exit Function
End If
Dim Cn$: Cn = FxDaoCnStr(Fx)
Dim Src$: Src = WsNm & "$"
DbtLnkFx = DbtLnk(A, T, Src, Cn)
End Function
Function TblLnkFb(TT, Fb$, Optional FbTny0) As String()
TblLnkFb = DbtLnkFb(CurrentDb, TT, Fb, FbTny0)
End Function
Function DbtLnkFb(A As Database, TT, Fb$, Optional FbTny0) As String()
Dim Tny$(), FbTny$()
Tny = CvNy(TT)
FbTny = CvNy(FbTny0)
    Select Case True
    Case Sz(FbTny) = Sz(Tny)
    Case Sz(FbTny) = 0
        FbTny = Tny
    Case Else
        Stop
    End Select
Dim Cn$: Cn = FbCnStr(Fb)
Dim J%, O$()
For J = 0 To UB(Tny)
    O = AyAdd(O, DbtLnk(A, Tny(J), FbTny(J), Cn))
Next
DbtLnkFb = O
End Function
Function TblLnkFx(T$, Fx$, Optional WsNm$ = "Sheet1") As String()
TblLnkFx = DbtLnkFx(CurrentDb, T, Fx, WsNm)
End Function
Function FbCnStr$(A)
FbCnStr = ";DATABASE=" & A & ";"
End Function

Function AyHas(A, M) As Boolean
Dim I
If Sz(A) = 0 Then Exit Function
For Each I In A
    If I = M Then
        AyHas = True
        Exit Function
    End If
Next
End Function

Function AyQuoteSqBkt(A) As String()
AyQuoteSqBkt = AyQuote(A, "[]")
End Function
Function DbtPk(A As Database, T) As String()

End Function
Function AyQuoteSng(A) As String()
AyQuoteSng = AyQuote(A, "'")
End Function

Function DbtStru$(A As Database, T$)
Dim Ay$()
Ay = DbtFny_zAutoQuote(A, T)
DbtStru = T & ": " & JnSpc(Ay)
End Function
Function DbttStru$(A As Database, TT)
Dim Tny$(), O$(), J%
Tny = CvNy(TT)
For J = 0 To UB(Tny)
    Push O, DbtStru(A, Tny(J))
Next
DbttStru = JnCrLf(O)
End Function
Sub DbtfChgDteToTxt(A As Database, T$, F)
A.Execute FmtQQ("Alter Table [?] add column [###] text(12)", T)
A.Execute FmtQQ("Update [?] set [###] = Format([?],'YYYY-MM-DD')", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [?]", T, F)
A.Execute FmtQQ("Alter Table [?] Add Column [?] text(12)", T, F)
A.Execute FmtQQ("Update [?] set [?] = [###]", T, F)
A.Execute FmtQQ("Alter Table [?] Drop Column [###]", T)
End Sub
Function JnComma$(A)
JnComma = Join(A, ",")
End Function
Function JnSpc$(A)
JnSpc = Join(A, " ")
End Function
Function UB&(A)
UB = Sz(A) - 1
End Function

Sub PushNonEmpty(O, A)
If A = "" Then Exit Sub
Push O, A
End Sub
Function DaoTy_Str$(T As DAO.DataTypeEnum)
Dim O$
Select Case T
Case DAO.DataTypeEnum.dbBoolean: O = "Boolean"
Case DAO.DataTypeEnum.dbDouble: O = "Double"
Case DAO.DataTypeEnum.dbText: O = "Text"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbByte: O = "Byte"
Case DAO.DataTypeEnum.dbInteger: O = "Int"
Case DAO.DataTypeEnum.dbLong: O = "Long"
Case DAO.DataTypeEnum.dbDouble: O = "Doubld"
Case DAO.DataTypeEnum.dbDate: O = "Date"
Case DAO.DataTypeEnum.dbDecimal: O = "Decimal"
Case DAO.DataTypeEnum.dbCurrency: O = "Currency"
Case DAO.DataTypeEnum.dbSingle: O = "Single"
Case Else: Stop
End Select
DaoTy_Str = O
End Function
Function DbqryRs(A As Database, Q) As DAO.Recordset
Set DbqryRs = A.QueryDefs(Q).OpenRecordset
End Function
Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function
Function Sz&(A)
On Error Resume Next
Sz = UBound(A) + 1
End Function
Function AyBrwEr(A) As Boolean
If Sz(A) = 0 Then Exit Function
AyBrwEr = True
AyBrw A
End Function
Sub AyBrw(A)
StrBrw Join(A, vbCrLf)
End Sub
Function TblFld_Ty(T, F) As DAO.DataTypeEnum
TblFld_Ty = CurrentDb.TableDefs(T).Fields(F).Type
End Function

Sub StrWrt(A, Ft$, Optional IsNotOvrWrt As Boolean)
Fso.CreateTextFile(Ft, Overwrite:=Not IsNotOvrWrt).Write A
End Sub
Sub FtBrw(A)
'Shell "code.cmd """ & A & """", vbHide
Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub
Function JnCrLf$(A)
JnCrLf = Join(A, vbCrLf)
End Function
Sub AyWrt(A, Ft$)
StrWrt JnCrLf(A), Ft
End Sub

Sub StrBrw(A)
Dim T$
T = TmpFt
StrWrt A, T
FtBrw T
End Sub
Function TmpFxm$(Optional Fdr$, Optional Fnn0$)
TmpFxm = TmpFfn(".xlsm", Fdr, Fnn0)
End Function

Function TmpFfn$(Optional Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
If Fnn0 = "" Then
    Fnn = TmpNm
Else
    Fnn = Fnn0
End If
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function
Function TmpFx$(Optional Fdr$, Optional Fnn$)
TmpFx = TmpFfn(".xlsx", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpHom & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function
Function DbtUpdToDteFld__1(A As Database, T$, KeyFld$, FmDteFld$) As Date()
Dim K$(), FmDte() As Date, ToDte() As Date, J&, CurKey$, NxtKey$, NxtFmDte As Date
With DbtRs(A, T)
    While Not .EOF
        Push FmDte, .Fields(FmDteFld).Value
        Push K, .Fields(KeyFld).Value
        .MoveNext
    Wend
End With
Dim U&
U = UB(K)
ReDim ToDte(U)
For J = 0 To U - 1
    CurKey = K(J)
    NxtKey = K(J + 1)
    NxtFmDte = FmDte(J + 1)
    If CurKey = NxtKey Then
        ToDte(J) = DateAdd("D", -1, NxtFmDte)
    Else
        ToDte(J) = DateSerial(2099, 12, 31)
    End If
Next
ToDte(U) = DateSerial(2099, 12, 31)
DbtUpdToDteFld__1 = ToDte
End Function
Sub ZZ_DbtUpdToDteFld()
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdToDteFld order by Sku,PermitDate"
DbtUpdToDteFld CurrentDb, "#A", "PermitDateEnd", "Sku", "PermitDate"
Stop
TblDrp "#A"
End Sub
Sub DbtUpdToDteFld(A As Database, T$, ToDteFld$, KeyFld$, FmDteFld$)
Dim ToDte() As Date, J&
ToDte = DbtUpdToDteFld__1(A, T, KeyFld, FmDteFld)
With DbtRs(A, T)
    While Not .EOF
        .Edit
        .Fields(ToDteFld).Value = ToDte(J): J = J + 1
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub

Function LinT1$(A)
LinT1 = LinShiftTerm(CStr(A))
End Function

Property Get TblImpSpec(T$, LnkSpec$, Optional WhBExpr$) As TblImpSpec
Dim O As New TblImpSpec
Set TblImpSpec = O.Init(T, LnkSpec$, WhBExpr)
End Property

Function TmpHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpHom = X
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function SqlDry(A$) As Variant()
SqlDry = DbqDry(CurrentDb, A)
End Function
Function DbqDry(A As Database, Sql$) As Variant()
Dim O()
With DbqRs(A, Sql)
    While Not .EOF
        Push O, FldsDr(.Fields)
        .MoveNext
    Wend
    .Close
End With
DbqDry = O
End Function
Function Xls(Optional Vis As Boolean) As Excel.Application
Static X As Boolean, Y As Excel.Application
Dim J%
Beg:
    J = J + 1
    If J > 10 Then Stop
If Not X Then
    X = True
    Set Y = New Excel.Application
End If
On Error GoTo xx
Dim A$
A = Y.Name
Set Xls = Y
If Vis Then XlsVis Y
Exit Function
xx:
    X = True
    GoTo Beg
End Function
Function DbtPutAtByCn(A As Database, T$, At As Range, Optional LoNm0$) As ListObject
If FstChr(T) <> "@" Then Stop
Dim LoNm$, Lo As ListObject
If LoNm0 = "" Then
    LoNm = "Tbl" & RmvFstChr(T)
Else
    LoNm = LoNm0
End If
Dim AtA1 As Range, CnStr, Ws As Worksheet
Set AtA1 = RgRC(At, 1, 1)
Set Ws = RgWs(At)
With Ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        FmtQQ("OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share D", A.Name) _
        , _
        "eny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Databa" _
        , _
        "se Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Je" _
        , _
        "t OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Com" _
        , _
        "pact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=" _
        , _
        "False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=AtA1).QueryTable '<---- At
        .CommandType = xlCmdTable
        .CommandText = Array(T) '<-----  T
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = LoNm '<------------ LoNm
        .Refresh BackgroundQuery:=False
    End With

End Function
Function NewXls(Optional Vis As Boolean) As Excel.Application
Dim O As New Excel.Application
If Vis Then O.Visible = True
Set NewXls = O
End Function
Function SqlStrCol(A) As String()
SqlStrCol = RsStrCol(CurrentDb.OpenRecordset(A))
End Function
Sub DicDmp(A As Dictionary)
Dim K
For Each K In A
    Debug.Print K, A(K)
Next
End Sub

Sub SqlAy_Run(SqlAy$())
Dim I
For Each I In SqlAy
    DoCmd.RunSQL I
Next
End Sub

Function RsStrCol(A As DAO.Recordset) As String()
Dim O$()
With A
    While Not .EOF
        Push O, .Fields(0).Value
        .MoveNext
    Wend
End With
RsStrCol = O
End Function
Function SqColInto(A, C%, OInto) As String()
Dim O
O = OInto
Erase O
Dim NR&, J&
NR = UBound(A, 1)
ReDim O(NR - 1)
For J = 1 To NR
    O(J - 1) = A(J, C%)
Next
SqColInto = O
End Function
Function SqColSy(A, C%) As String()
SqColSy = SqColInto(A, C, EmpSy)
End Function
Function AtVBar(A As Range) As Range
If IsEmpty(A.Value) Then Stop
If IsEmpty(RgRC(A, 2, 1).Value) Then
    Set AtVBar = RgRC(A, 1, 1)
    Exit Function
End If
Set AtVBar = RgCRR(A, 1, 1, A.End(xlDown).Row - A.Row + 1)
End Function
Function RgCRR(A As Range, C, R1, R2) As Range
Set RgCRR = RgRCRC(A, R1, C, R2, C)
End Function
Function SqSyV(A) As String()
SqSyV = SqColSy(A, 1)
End Function
Sub RgFillCol(A As Range)
Dim Rg As Range
Dim Sq()
Sq = SqzVBar(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub
Sub RgFillRow(A As Range)
Dim Rg As Range
Dim Sq()
Sq = SqzHBar(A.Rows.Count)
RgReSz(A, Sq).Value = Sq
End Sub
Function SqzVBar(N%) As Variant()
Dim O(), J%
ReDim O(1 To N, 1 To 1)
For J = 1 To N
    O(J, 1) = J
Next
SqzVBar = O
End Function
Function SqzHBar(N%) As Variant()
Dim O(), J%
ReDim O(1 To 1, 1 To N)
For J = 1 To N
    O(1, J) = J
Next
SqzHBar = O
End Function
Sub FxOpn(A$)
If Not FfnIsExist(A) Then
    MsgBox "File not found: " & vbCrLf & vbCrLf & A
    Exit Sub
End If
Dim C$
C = FmtQQ("Excel ""?""", A)
Debug.Print C
Shell C, vbMaximizedFocus
'Xls(Vis:=True).Workbooks.Open A
End Sub
Function AyQuote(A, Q$) As String()
If Sz(A) = 0 Then Exit Function
Dim Q1$, Q2$
Select Case True
Case Len(Q) = 1: Q1 = Q: Q2 = Q
Case Len(Q) = 2: Q1 = Left(Q, 1): Q2 = Right(Q, 1)
Case Else: Stop
End Select

Dim I, O$()
For Each I In A
    Push O, Q1 & I & Q2
Next
AyQuote = O
End Function
Function FldsDr(A As DAO.Fields) As Variant()
Dim O(), F As DAO.Field
For Each F In A
    Push O, F.Value
Next
FldsDr = O
End Function
Function SubStrCnt%(A, SubStr$)
Dim J&, O%, P%, L%
L = Len(SubStr)
P = InStr(A, SubStr)
While P > 0
    O = O + 1
    J = J + 1: If J > 100000 Then Stop
    P = InStr(P + L, A, SubStr)
Wend
SubStrCnt = O
End Function
Function RgCC(A As Range, C1, C2) As Range
Set RgCC = RgRCRC(A, 1, C1, A.Rows.Count, C2)
End Function

Sub ZZ_FmtQQAv()
Debug.Print FmtQQ("klsdf?sdf?dsklf", 2, 1)
End Sub
Function FmtQQAv$(QQVbl, Av)
Dim O$, I, Cnt
O = Replace(QQVbl, "|", vbCrLf)
Cnt = SubStrCnt(QQVbl, "?")
If Cnt <> Sz(Av) Then Stop
For Each I In Av
    O = Replace(O, "?", I, Count:=1)
Next
FmtQQAv = O
End Function
Sub PushAy(O, A)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Push O, I
Next
End Sub

Function AyIsEmpty(A) As Boolean
AyIsEmpty = Sz(A) = 0
End Function
Function AyIsAllEq(A) As Boolean
If Sz(A) <= 1 Then AyIsAllEq = True: Exit Function
Dim A0, J&
A0 = A(0)
For J = 2 To UB(A)
    If A0 <> A(0) Then Exit Function
Next
AyIsAllEq = True
End Function
Function FfnNxt$(A$)
If Not FfnIsExist(A) Then FfnNxt = A: Exit Function
Dim J%, O$
For J = 1 To 99
    O = FfnNxtN(A, J)
    If Not FfnIsExist(O) Then FfnNxt = O: Exit Function
Next
Stop
End Function

Function FfnAddFnSfx$(A$, Sfx$)
FfnAddFnSfx = FfnPth(A) & FfnFnn(A) & Sfx & FfnExt(A)
End Function

Function FfnNxtN$(A$, N%)
If 1 > N Or N > 99 Then Stop
Dim Sfx$
Sfx = "(" & Format(N, "00") & ")"
FfnNxtN = FfnAddFnSfx(A, Sfx)
End Function

Function PthSel$(A, Optional Tit$ = "Select a Path", Optional BtnNm$ = "Use this path")
With FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .InitialFileName = Nz(A, "")
    .Show
    If .SelectedItems.Count = 1 Then
        PthSel = PthEnsSfx(.SelectedItems(1))
    End If
End With
End Function
Sub ZZ_PthSel()
MsgBox FfnSel("C:\")
End Sub
Function FfnSel$(A, Optional FSpec$ = "*.*", Optional Tit$ = "Select a file", Optional BtnNm$ = "Use the File Name")
With FileDialog(msoFileDialogFilePicker)
    .Filters.Clear
    .Title = Tit
    .AllowMultiSelect = False
    .Filters.Add "", FSpec
    .InitialFileName = A
    .ButtonName = BtnNm
    .Show
    If .SelectedItems.Count = 1 Then
        FfnSel = .SelectedItems(1)
    End If
End With
End Function
Sub TxtbSelPth(A As Access.TextBox)
Dim R$
R = PthSel(A.Value)
If R = "" Then Exit Sub
A.Value = R
End Sub
Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnCutExt(FfnFn(A))
End Function
Function FfnCutExt$(A)
Dim B$, C$, P%
B = FfnFn(A)
P = InStrRev(B, ".")
If P = 0 Then
    C = B
Else
    C = Left(B, P - 1)
End If
FfnCutExt = FfnPth(A) & C
End Function
Function PthEns$(A$)
If Dir(A, VbFileAttribute.vbDirectory) = "" Then MkDir A
PthEns = A
End Function

Function PthFfnAy(A, Optional Spec$ = "*.*") As String()
Dim O$(), B$, P$
P = PthEnsSfx(A)
B = Dir(A & Spec)
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    Push O, P & B
    B = Dir
Wend
PthFfnAy = O
End Function

Function FfnExt$(Ffn)
Dim P%: P = InStrRev(Ffn, ".")
If P = 0 Then Exit Function
FfnExt = Mid(Ffn, P)
End Function

Function PthFxAy(A) As String()
Dim O$(), B$
If Right(A, 1) <> "\" Then Stop
B = Dir(A & "*.xls")
Dim J%
While B <> ""
    J = J + 1
    If J > 1000 Then Stop
    If FfnExt(B) = ".xls" Then
        Push O, A & B
    End If
    B = Dir
Wend
PthFxAy = O
End Function

Function RmvLasChr$(A)
RmvLasChr = Left(A, Len(A) - 1)
End Function
Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function AyIsEq(A, B) As Boolean
Dim U&, J&
U = UB(A)
If UB(B) <> U Then Exit Function
For J = 0 To U
    If A(J) <> B(J) Then Exit Function
Next
AyIsEq = True
End Function
Function RsIsBrk(A As DAO.Recordset, GpKy$(), LasVy()) As Boolean
RsIsBrk = Not AyIsEq(RsVy(A, GpKy), LasVy)
End Function
Function RsVy(A As DAO.Recordset, Optional Ky0) As Variant()
RsVy = FldsVy(A.Fields, Ky0)
End Function
Function FldsVyByKy(A As DAO.Fields, Ky$()) As Variant()
Dim O(), J%, K
If Sz(Ky) = 0 Then
    FldsVyByKy = ItrVy(A)
    Exit Function
End If
ReDim O(UB(Ky))
For Each K In Ky
    O(J) = A(K).Value
    J = J + 1
Next
FldsVyByKy = O
End Function
Sub ZZ_FldsVy()
Dim Rs As DAO.Recordset, Vy()
Set Rs = CurrentDb.OpenRecordset("Select * from SkuB")
With Rs
    While Not .EOF
        Vy = RsVy(Rs)
        Debug.Print JnComma(Vy)
        .MoveNext
    Wend
    .Close
End With
End Sub
Function ItrPrpAy(A, PrpNm$) As Variant()
Dim O(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
ItrPrpAy = O
End Function
Function ItrVy(A) As Variant()
ItrVy = ItrPrpAy(A, "Value")
End Function
Function IsDte(A) As Boolean
IsDte = VarType(A) = vbDate
End Function
Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function
Function IsSy(A) As Boolean
IsSy = VarType(A) = vbString + vbArray
End Function
Function CvSy(A) As String()
CvSy = A
End Function
Function FldsVy(A As DAO.Fields, Optional Ky0) As Variant()
Select Case True
Case IsMissing(Ky0)
    FldsVy = ItrVy(A)
Case IsStr(Ky0)
    FldsVy = FldsVyByKy(A, SslSy(Ky0))
Case IsSy(Ky0)
    FldsVy = FldsVyByKy(A, CvSy(Ky0))
Case Else
    Stop
End Select
End Function
Private Sub ZZ_SslSqBktCsv()
Debug.Print SslSqBktCsv("a b c")
End Sub
Function SslSqBktCsv$(A)
Dim B$(), C$()
B = SslSy(A)
C = AyQuoteSqBkt(B)
SslSqBktCsv = JnComma(C)
End Function
Function Ny0SqBktCsv$(A)
Dim B$(), C$()
B = CvNy(A)
C = AyQuoteSqBkt(B)
Ny0SqBktCsv = JnComma(C)
End Function
Function RsFny(A As DAO.Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function AyHasAy(A, Ay) As Boolean
Dim I
For Each I In Ay
    If Not AyHas(A, I) Then Exit Function
Next
AyHasAy = True
End Function

Function SqlQQStr_Sy(Sql$, QQStr$) As String()
Dim Dry: Dry = SqlDry(Sql)
If AyIsEmpty(Dry) Then Exit Function
Dim O$()
Dim Dr
For Each Dr In Dry
    Push O, FmtQQAv(QQStr, Dr)
Next
SqlQQStr_Sy = O
End Function


Function FldsCsv$(A As DAO.Fields)
FldsCsv = AyCsv(ItrVy(A))
End Function
Function VarCsv$(A)
Select Case True
Case IsStr(A): VarCsv = """" & A & """"
Case IsDte(A): VarCsv = Format(A, "YYYY-MM-DD HH:MM:SS")
Case Else: VarCsv = Nz(A, "")
End Select
End Function
Function AyMapInto(A, MapFunNm$, OInto)
Dim J&, O, I, U&
O = OInto
Erase O
U = UB(A)
If U = -1 Then
    AyMapInto = O
    Exit Function
End If
ReDim O(U)
For Each I In A
    Asg Run(MapFunNm, I), O(J)
    J = J + 1
Next
AyMapInto = O
End Function
Sub Asg(Fm, OTo)
If IsObject(Fm) Then
    Set OTo = Fm
Else
    OTo = Nz(Fm, "")
End If
End Sub
Function AyMapSy(A, MapFunNm$) As String()
AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
End Function
Function AyCsv$(A)
AyCsv = Join(A, ",")
Exit Function
Dim J%
For J = 0 To UB(A)
    A(J) = VarCsv(A(J))
Next
AyCsv = Join(A, ",")
End Function
Sub ZZ_DbtUpdSeq()
DoCmd.SetWarnings False
DoCmd.RunSQL "Select * into [#A] from ZZ_DbtUpdSeq order by Sku,PermitDate"
DoCmd.RunSQL "Update [#A] set BchRateSeq=0, Rate=Round(Rate,0)"
DbtUpdSeq CurrentDb, "#A", "BchRateSeq", "Sku", "Sku Rate"
TblOpn "#A"
Stop
DoCmd.RunSQL "Drop Table [#A]"
End Sub

Sub DbtUpdSeq(A As Database, T$, SeqFldNm$, Optional RestFny0, Optional IncFny0)
'Assume T is sorted
'
'Update A->T->SeqFldNm using RestFny0,IncFny0, assume the table has been sorted
'Update A->T->SeqFldNm using OrdFny0, RestFny0,IncFny0
Dim RestFny$(), IncFny$(), Sql$
Dim LasRestVy(), LasIncVy(), Seq&, OrdS$, Rs As DAO.Recordset
'OrdFny RestAy IncAy Sql
RestFny = CvNy(RestFny0)
IncFny = CvNy(IncFny0)
If Sz(RestFny) = 0 And Sz(IncFny) = 0 Then
    With A.OpenRecordset(T)
        Seq = 1
        While Not .EOF
            .Edit
            .Fields(SeqFldNm) = Seq
            Seq = Seq + 1
            .Update
            .MoveNext
        Wend
        .Close
    End With
    Exit Sub
End If
'--
Set Rs = A.OpenRecordset(T) ', RecordOpenOptionsEnum.dbOpenForwardOnly, dbForwardOnly)
With Rs
    While Not .EOF
        If RsIsBrk(Rs, RestFny, LasRestVy) Then
            Seq = 1
            LasRestVy = RsVy(Rs, RestFny)
            LasIncVy = RsVy(Rs, IncFny)
        Else
            If RsIsBrk(Rs, IncFny, LasIncVy) Then
                Seq = Seq + 1
                LasIncVy = RsVy(Rs, IncFny)
            End If
        End If
        .Edit
        .Fields(SeqFldNm).Value = Seq
        .Update
        .MoveNext
    Wend
End With
End Sub

Function CvRg(A) As Range
Set CvRg = A
End Function

Function PnmFn$(A$)
PnmFn = PnmVal(A & "Fn")
End Function

Function RsCsv$(A As DAO.Recordset)
RsCsv = FldsCsv(A.Fields)
End Function

Function AyQuoteSqBktCsv$(A)
AyQuoteSqBktCsv = JnComma(AyQuoteSqBkt(A))
End Function

Function LinRmvTerm$(ByVal A$)
LinShiftTerm A
LinRmvTerm = A
End Function
Sub AppExp()
PthClr SrcPth
AppExpMd
AppExpFrm
AppExpStru
End Sub
Sub AppExpFrm()
Dim Nm$, P$, I
P = SrcPth
For Each I In AppFrmNy
    Nm = I
    SaveAsText acForm, Nm, P & Nm & ".Frm.Txt"
Next
End Sub
Function AppFrmNy() As String()
AppFrmNy = ItrNy(CodeProject.AllForms)
End Function
Function AppMdNy() As String()
AppMdNy = ItrNy(CodeProject.AllModules)
End Function
Function Stru$()
Stru = DbStru(CurrentDb)
End Function
Sub AppExpStru()
StruWrt Stru, SrcPth & "Stru.txt"
End Sub
Sub FfnDltIfExist(A$)
On Error GoTo X
If FfnIsExist(A) Then Kill A
Exit Sub
X:
Debug.Print "FfnDltIfExist: Unable to delete file [" & A & "].  Er[" & Err.Description & "]"
End Sub
Sub FfnAy_DltIfExist(A)
AyDo A, "FfnDltIfExist"
End Sub

Sub PthClr(A$)
FfnAy_DltIfExist PthFfnAy(A)
End Sub
Function SrcPth$()
Dim X As Boolean, Y$
If Not X Then
    X = True
    Y = CurDbPth & "Src\"
    PthEns Y
End If
Y = SrcPth
End Function
Sub AppExpMd()
Dim MdNm$, I, P$
P = SrcPth
For Each I In AppMdNy
    MdNm = I
    SaveAsText acModule, MdNm, P & MdNm & ".bas"
Next
End Sub
Sub ZZ_DbtReSeqFld()
DbtReSeqFld CurrentDb, "ZZ_DbtUpdSeq", "Permit PermitD"
End Sub

Sub DbtReSeqFld(A As Database, T$, ReSeqSpec$)
DbtReSeqFldByFny A, T, ReSeqSpec_Fny(ReSeqSpec)
End Sub

Function DiczTRLy(TermRestLy$()) As Dictionary
Dim I, L$, K$, O As New Dictionary
If Sz(TermRestLy) > 0 Then
    For Each I In TermRestLy
        L = I
        K = LinShiftTerm(L)
        O.Add K, L
    Next
End If
Set DiczTRLy = O
End Function

Sub ZZ_ReSeqSpec_Fny()
AyBrw ReSeqSpec_Fny("Flg RecTy Amt Key Uom MovTy Qty BchRateUX RateTy Bch Las GL |" & _
" Flg IsAlert IsWithSku |" & _
" Key Sku PstMth PstDte |" & _
" Bch BchNo BchPermitDate BchPermit |" & _
" Las LasBchNo LasPermitDate LasPermit |" & _
" GL GLDocNo GLDocDte GLAsg GLDocTy GLLin GLPstKy GLPc GLAc GLBusA GLRef |" & _
" Uom Des StkUom Ac_U")
End Sub

Function ReSeqSpec_Fny(A$) As String()
Dim Ay$(), D As Dictionary, O$(), L1$, L
Ay = SplitVBar(A)
If Sz(Ay) = 0 Then Exit Function
L1 = AyShift(Ay)
Set D = DiczTRLy(Ay)
For Each L In SslSy(L1)
    If D.Exists(L) Then
        Push O, D(L)
    Else
        Push O, L
    End If
Next
ReSeqSpec_Fny = SslSy(JnSpc(O))
End Function
Sub DbReOpn(A As Database)
Dim Nm$
Nm = A.Name
A.Close
Set A = DAO.DBEngine.OpenDatabase(Nm)
End Sub
Sub DbtReSeqFldByFny(A As Database, T$, Fny$())
Dim TFny$(), F$(), J%, FF
TFny = DbtFny(A, T)
If Sz(TFny) = Sz(Fny) Then
    F = Fny
Else
    F = AyAdd(Fny, AyMinus(TFny, Fny))
End If
For Each FF In F
    J = J + 1
    A.TableDefs(T).Fields(FF).OrdinalPosition = J
Next
End Sub
Function OyDrs(A, PrpNy0) As Drs
Dim Fny$(), Dry()
Fny = CvNy(PrpNy0)
Dry = OyDry(A, Fny)
Set OyDrs = Drs(Fny, Dry)
End Function
Function ObjDr(A, PrpNy0) As Variant()
Dim PrpNy$(), U%, O(), J%
PrpNy = CvNy(PrpNy0)
U = UB(PrpNy)
ReDim O(U)
For J = 0 To U
    Asg ObjPrp(A, PrpNy(J)), O(J)
Next
ObjDr = O
End Function
Function OyDry(A, PrpNy0) As Variant()
Dim O(), U%, I
Dim PrpNy$()
PrpNy = CvNy(PrpNy0)
For Each I In A
    Push O, ObjDr(I, PrpNy)
Next
OyDry = O
End Function
Sub ZZ_OyDrs()
WsVis DrsWs(OyDrs(CurrentDb.TableDefs("ZZ_DbtUpdSeq").Fields, "Name Type OrdinalPosition"))
End Sub
Function DrsWs(A As Drs) As Worksheet
Set DrsWs = SqWs(DrsSq(A))
End Function
Function DrsPutAt(A As Drs, At As Range) As Range
Set DrsPutAt = SqPutAt(DrsSq(A), At)
End Function
Function DryWs(A) As Worksheet
Set DryWs = SqWs(DrySq(A))
End Function
Function DryNCol%(A)
Dim O%, Dr
For Each Dr In A
    O = Max(O, Sz(Dr))
Next
DryNCol = O
End Function
Function DrySq(A) As Variant()
Dim O(), C%, R&, Dr
Dim NC%, NR&
NC = DryNCol(A)
NR = Sz(A)
ReDim O(1 To NR, 1 To NC)
For R = 1 To NR
    Dr = A(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R, C) = Dr(C - 1)
    Next
Next
DrySq = O
End Function
Function DbPth$(A As Database)
DbPth = FfnPth(A.Name)
End Function
Function DrsNCol%(A As Drs)
DrsNCol = Max(Sz(A.Fny), DryNCol(A.Dry))
End Function
Sub TpWrtFfn(Ffn$)
AttExp "Tp", Ffn
End Sub
Sub AAA()
Dim A As New Excel.Application
A.Visible = True
Stop
End Sub
Sub TpExp()
AttExp "Tp", TpFx
End Sub
Sub TpImp()
Dim A$
Const Trc As Boolean = True
A = TpFx
If Not FfnIsExist(A) Then
    If Trc Then
        Debug.Print "-----"
        Debug.Print "TpImp"
        Debug.Print "Given-Tp   : "; A
        Debug.Print "Given-Tp is: Not exist"
    End If
End If
If AttIsOld("Tp", A) Then AttImp "Tp", A
End Sub
Function AttIsOld(A$, Ffn$) As Boolean
Dim T1 As Date, T2 As Date
T1 = AttTim(A)
T2 = FfnTim(Ffn)
If True Then
    Debug.Print "---------"
    Debug.Print "AttIsOld:"
    Debug.Print "Att         : "; A
    Debug.Print "GiveFfn     : "; Ffn
    Debug.Print "AttTim      : "; T1
    Debug.Print "GiveFfn-Tim : "; T2
    Debug.Print "AttIsOld    : "; T1 < T2
End If
AttIsOld = T1 < T2
End Function
Function TblkfV(T$, K$, F$)
TblkfV = DbtkfV(CurrentDb, T, K, F)
End Function
Function DbtPkNm$(A As Database, T$)
Dim I As DAO.Index
For Each I In A.TableDefs(T).Indexes
    If I.Primary Then DbtPkNm = I.Name
Next
End Function
Function DbtkfV(A As Database, T$, K$, F$)
With A.TableDefs(T).OpenRecordset
    .Index = DbtPkNm(A, T)
    .Seek "=", K
    If Not .EOF Then DbtkfV = .Fields(F)
End With
End Function
Function AttTim(A$) As Date
AttTim = TblkfV("Att", A, "FilTim")
End Function
Function AttSz(A$) As Date
AttSz = TblkfV("Att", A, "FilSz")
End Function
Function DrsSq(A As Drs) As Variant()
Dim O(), C%, R&, Dr(), Dry()
Dim Fny$(), NC%, NR&
Dry = A.Dry
Fny = A.Fny

NR = Sz(Dry)
NC = DrsNCol(A)
If Sz(Fny) <> NC Then Stop
ReDim O(1 To NR + 1, 1 To NC)
For C = 1 To NC
    O(1, C) = Fny(C - 1)
Next
For R = 1 To NR
    Dr = Dry(R - 1)
    For C = 1 To Min(Sz(Dr), NC)
        O(R + 1, C) = Dr(C - 1)
    Next
Next
DrsSq = O
End Function
Function SqLo(A, Optional LoNm$) As ListObject
Set SqLo = SqAt_Lo(A, NewA1, LoNm)
End Function
Function SqWs(A) As Worksheet
Set SqWs = LoWs(SqLo(A))
End Function
Sub WImpTbl(TT)
DbtImpTbl W, TT
End Sub

Function WbMax(A As Workbook) As Workbook
A.Application.WindowState = xlMaximized
Set WbMax = A
End Function

Function WtChkCol(T$, LnkColStr$) As String()
WtChkCol = DbtChkCol(W, T, LnkColStr)
End Function

Function WtLnkFb(T$, Fb$) As String()
WtLnkFb = DbtLnkFb(W, T, Fb)
End Function

Sub WQuit()
WCls
Quit
End Sub

Function WtLnkFx(T$, Fx$, Optional WsNm$ = "Sheet1") As String()
WtLnkFx = DbtLnkFx(W, T, Fx, WsNm)
End Function

Sub WOpn()
FbEns WFb
Set W = FbDb(WFb)
End Sub

Sub WIni()
TpImp
WCls
FfnDltIfExist WFb
WOpn
End Sub

Sub WRun(A$)
On Error GoTo X
W.Execute A
Exit Sub
X:
Debug.Print A
Debug.Print "?WStru("""")"
DbCrtQry W, "Query1", A
Stop
End Sub