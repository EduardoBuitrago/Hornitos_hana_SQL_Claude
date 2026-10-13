SELECT
     'N' AS 'Sel',
      Name AS 'ItemCode',
      U_TipDoc AS 'Tipo doc',
      U_NumDoc AS 'Num Doc',
      U_NumLot AS 'Numero lote',
      U_CantLote AS 'Cantidad disponible',
	  CAST(0 AS DECIMAL(19,6)) AS 'Cantidad',
      U_FchVenLt AS 'Fecha vence',
      U_FchAdmLt AS 'Fecha Admin',
      U_almacen AS 'Almacen'
  FROM [@CLC_LTS] T0
  WHERE 
  T0.Name = '{0}'
  and T0.U_CantLote > 0
  AND U_almacen = '{1}'
  order by U_FchVenLt