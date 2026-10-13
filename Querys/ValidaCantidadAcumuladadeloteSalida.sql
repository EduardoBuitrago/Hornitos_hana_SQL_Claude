  SELECT
      Name AS 'ItemCode',
      SUM(U_CantLote) AS 'Cantidad',
      U_almacen AS 'Almacen'
  FROM [@CLC_LTS] T0
  WHERE 
  T0.Name = '{0}'
  and T0.U_CantLote > 0
  AND U_almacen = '{1}'
  GROUP BY Name,U_almacen
  HAVING SUM(U_CantLote) >= {2}