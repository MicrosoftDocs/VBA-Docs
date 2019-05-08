---
title: PivotTable.EnableWriteback property (Excel)
keywords: vbaxl10.chm235186
f1_keywords:
- vbaxl10.chm235186
ms.prod: excel
api_name:
- Excel.PivotTable.EnableWriteback
ms.assetid: d13b3db8-070a-3b29-9ff7-bfdcd143e5fa
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.EnableWriteback property (Excel)

Returns or sets whether writing back to the data source is enabled for the specified PivotTable. The default value is **False**. Read/write.


## Syntax

_expression_.**EnableWriteback**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Return value

**Boolean**


## Remarks

For OLAP data sources, setting the **EnableWriteback** property to **True** enables writeback and disables the alert for when the user overwrites values in the data area of the PivotTable. For non-OLAP data sources, setting the **EnableWriteback** property to **True** enables write-back in code, and also allows the user to change data values that previously could not be changed.

The **EnableWriteback** and **[EnableDataValueEditing](Excel.PivotTable.EnableDataValueEditing.md)** properties cannot be set to **True** at the same time.

If the **EnableDataValueEditing** property is set to **True** and then the **EnableWriteback** property is set to **True**, the **EnableDataValueEditing** property is set to **False** automatically, the PivotTable is refreshed, and any editing performed on data values is lost.

If the **EnableWriteback** property is set to **True** and then the **EnableDataValueEditing** property is set to **True**, the **EnableWriteback** property is set to **False** automatically, the PivotTable is not refreshed, and the data source values are restored.

For non-OLAP data sources, setting this property generates a run-time error.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]