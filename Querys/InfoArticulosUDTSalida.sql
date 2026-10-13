SELECT
      Name AS 'ItemCode',
      U_TipDoc AS 'Tipo doc',
      U_NumDoc AS 'Num Doc',
      U_NumLot AS 'Numero lote',
      U_Cantidad AS 'Cantidad',
      U_FchSal AS 'Fecha salida',
      U_almacen AS 'Almacen'
  FROM [@CLC_LTSSA] T0
  WHERE 
  T0.Name = '{0}' AND 
  T0.U_NumDoc = {1} AND 
  T0.U_TipDoc = '{2}'