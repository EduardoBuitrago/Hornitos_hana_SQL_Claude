UPDATE 
[@CLC_LTS]
SET U_CantLote = (U_CantLote - {0})
WHERE 
Name = '{1}' AND
U_NumLot = '{4}' AND
U_almacen = '{5}'