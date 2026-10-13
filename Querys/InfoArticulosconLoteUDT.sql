SELECT 
      Name AS 'ItemCode',
      U_TipDoc AS 'Tipo doc',
      U_NumDoc AS 'Num Doc',
      U_NumLot AS 'Numero lote',
      U_CantLote AS 'Cantidad',
      U_FchVenLt AS 'Fecha vence',
      U_FchAdmLt AS 'Fecha Admin',
      U_almacen AS 'Almacen'
  FROM [@CLC_LTS] T0
  WHERE 
  T0.Name = '{0}'
  AND T0.U_TipDoc = '{1}'
  AND T0.U_NumDoc = {2}