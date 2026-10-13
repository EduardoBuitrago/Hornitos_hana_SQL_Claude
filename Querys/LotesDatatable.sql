SELECT 
'N' AS 'Sel',
'{0}' AS 'ItemCode',
{1} AS 'Num Doc',
'{2}' AS 'Tipo doc',
CAST(({3}) AS DECIMAL(19,6)) AS 'Cantidad',
CAST('{4}' AS NVARCHAR(50)) AS 'Numero lote',
CONVERT(DATETIME,'{5}',103) AS 'Fecha vence',
CONVERT(DATETIME,'{6}',103) AS 'Fecha Admin',
'{7}' AS 'Almacen',
{8} AS 'Index',
{9} AS 'NumLinea'