SELECT
T0.U_NumLot,
T0.U_Cantidad
FROM [@CLC_LTSSA] T0
WHERE 
T0.Name = '{0}' AND 
T0.U_NumDoc = {1} AND 
T0.U_TipDoc = '{2}'

