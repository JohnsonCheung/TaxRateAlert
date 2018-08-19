Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Public Nm$, Ty As DAO.DataTypeEnum, Extnm$
Friend Property Get Init(Nm, Ty As DAO.DataTypeEnum, Extnm$)
Me.Nm = Nm
Me.Ty = Ty
Me.Extnm = Extnm
Set Init = Me
End Property