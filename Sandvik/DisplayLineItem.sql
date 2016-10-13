use entssi
go
SELECT PurchaseOrder.POPK,
PurchaseOrder.Total,
PurchaseOrderDetail.PartName,
PurchaseOrderDetail.PartID,
PurchaseOrderDetail.LineItemNo,
PurchaseOrderReceive.OrderUnitQtyReceived,
PurchaseOrderDetail.OrderUnitQty
from PurchaseOrder 
inner join PurchaseOrderDetail on PurchaseOrderDetail.POPK= PurchaseOrder .POPK
inner join PurchaseOrderReceive on PurchaseOrderReceive.POPK =PurchaseOrder.POPK
order by popk asc