Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Public Tbl$, LnkColVbl$, WhBExpr$
Friend Property Get Init(Tbl$, LnkColVbl$, Optional WhBExpr$)
Me.Tbl = Tbl
Me.LnkColVbl = LnkColVbl
Me.WhBExpr = WhBExpr
Set Init = Me
End Property