SELECT
'{0}' AS 'ItemCode',
'{1}' AS 'Tipo doc',
{2} AS 'Num Doc',
'{3}' AS 'Numero lote',
{4} AS 'Cantidad',
CONVERT(DATETIME,'{5}',103) AS 'Fecha vence',
CONVERT(DATETIME,'{6}',103) AS 'Fecha Admin',
'{7}' AS 'Almacen',
'{8}' AS 'Index',
{9} AS 'NumLinea'
