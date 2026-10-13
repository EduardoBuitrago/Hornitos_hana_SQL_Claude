SELECT
'N' AS 'Sel',
'{0}' AS 'ItemCode',
'{1}' AS 'Tipo doc',
{2} AS 'Num Doc',
CAST('' AS NVARCHAR(50)) AS 'Numero lote',
CAST(({3}) AS DECIMAL(19,6)) AS 'Cantidad',
ISNULL((SELECT DATEADD(DAY,T0.U_V_Util,GETDATE()) FROM OITM T0 WHERE T0.ItemCode = '{0}'),'') AS 'Fecha vence',
GETDATE() AS 'Fecha Admin',
'{4}' as 'Almacen',
'' AS 'Index',
{5} AS 'NumLinea'