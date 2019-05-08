---
title: PivotTable.EnableDrilldown property (Excel)
keywords: vbaxl10.chm235106
f1_keywords:
- vbaxl10.chm235106
ms.prod: excel
api_name:
- Excel.PivotTable.EnableDrilldown
ms.assetid: 329e6c74-6b23-eac8-2ffb-45696076c712
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.EnableDrilldown property (Excel)

**True** if drilldown is enabled. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableDrilldown**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Setting this property for a PivotTable report sets it for all fields in that report.

For OLAP data sources, the value is always **True**.


## Example

This example disables drilldown for all fields in the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables("Pivot1").EnableDrilldown = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]