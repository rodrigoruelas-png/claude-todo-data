-- PROPUESTA — NO EJECUTADA NI PROBADA. Solo SELECT. Usar únicamente en ambiente de
-- DESARROLLO/PRUEBA con usuario de solo lectura. Nunca en producción sin tu aprobación.
-- Parámetros: :org_code (ej. 'MC1'), :dias_sin_mov (ej. 180)
-- Verificar nombres de columnas/tablas contra tu versión de EBS antes de usar.

-- 1) Existencia por P/N con antigüedad (fecha de recepción) y min/max
SELECT msi.segment1                        AS pn,
       msi.description,
       SUM(moq.transaction_quantity)       AS existencia,
       MIN(moq.orig_date_received)         AS primera_recepcion,
       TRUNC(SYSDATE - MIN(moq.orig_date_received)) AS dias_mas_antiguo,
       msi.min_minmax_quantity             AS minimo,
       msi.max_minmax_quantity             AS maximo
FROM   mtl_onhand_quantities_detail moq
JOIN   mtl_parameters      mp  ON mp.organization_id = moq.organization_id
JOIN   mtl_system_items_b  msi ON msi.inventory_item_id = moq.inventory_item_id
                              AND msi.organization_id   = moq.organization_id
WHERE  mp.organization_code = :org_code
GROUP  BY msi.segment1, msi.description, msi.min_minmax_quantity, msi.max_minmax_quantity;

-- 2) P/N con existencia y SIN movimiento en :dias_sin_mov días
WITH ult AS (
  SELECT mmt.inventory_item_id, MAX(mmt.transaction_date) AS ult_mov
  FROM   mtl_material_transactions mmt
  JOIN   mtl_parameters mp ON mp.organization_id = mmt.organization_id
  WHERE  mp.organization_code = :org_code
  GROUP  BY mmt.inventory_item_id
), oh AS (
  SELECT moq.inventory_item_id, SUM(moq.transaction_quantity) AS existencia
  FROM   mtl_onhand_quantities_detail moq
  JOIN   mtl_parameters mp ON mp.organization_id = moq.organization_id
  WHERE  mp.organization_code = :org_code
  GROUP  BY moq.inventory_item_id
)
SELECT msi.segment1 AS pn, oh.existencia, ult.ult_mov
FROM   oh
JOIN   mtl_parameters mp ON mp.organization_code = :org_code
JOIN   mtl_system_items_b msi ON msi.inventory_item_id = oh.inventory_item_id
                             AND msi.organization_id   = mp.organization_id
LEFT   JOIN ult ON ult.inventory_item_id = oh.inventory_item_id
WHERE  oh.existencia > 0
AND   (ult.ult_mov IS NULL OR ult.ult_mov < SYSDATE - :dias_sin_mov);

-- 3) Ítems críticos: requerimiento abierto de WIP (jobs liberados, status 3) mayor que la existencia
SELECT msi.segment1 AS pn,
       SUM(wro.required_quantity - wro.quantity_issued) AS requerido_pendiente,
       NVL(MAX(oh.existencia), 0)                       AS existencia
FROM   wip_requirement_operations wro
JOIN   wip_discrete_jobs wdj ON wdj.wip_entity_id = wro.wip_entity_id
                            AND wdj.organization_id = wro.organization_id
JOIN   mtl_parameters mp ON mp.organization_id = wro.organization_id
JOIN   mtl_system_items_b msi ON msi.inventory_item_id = wro.inventory_item_id
                             AND msi.organization_id   = wro.organization_id
LEFT   JOIN (SELECT moq.inventory_item_id, moq.organization_id,
                    SUM(moq.transaction_quantity) AS existencia
             FROM   mtl_onhand_quantities_detail moq
             GROUP  BY moq.inventory_item_id, moq.organization_id) oh
       ON oh.inventory_item_id = wro.inventory_item_id
      AND oh.organization_id   = wro.organization_id
WHERE  mp.organization_code = :org_code
AND    wdj.status_type = 3
GROUP  BY msi.segment1
HAVING SUM(wro.required_quantity - wro.quantity_issued) > NVL(MAX(oh.existencia), 0);
