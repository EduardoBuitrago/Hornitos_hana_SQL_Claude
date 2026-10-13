select 
T0.Name
,'{3}' as  'TipDoc'
,{1} as 'NumDoc'
,T0.U_NumDoc as 'NumDoc2'
,T0.U_NumLot
,ABS((T0.U_CantLote - {6})) as 'U_CantLote'
,T0.U_FchVenLt
,T0.U_FchAdmLt
,T0.U_almacen
,T0.U_TipoCont
,T0.U_TpDocEnt
,T0.U_NmDocEnt
,T0.U_RowNum
from [@CLC_LTSSATEMP] T0
where
T0.Name = '{0}' and
T0.U_NumDoc = {5} and 
T0.U_almacen = '{2}' and 
T0.U_RowNum = {4} and 
(T0.U_CantLote - {6}) != 0
