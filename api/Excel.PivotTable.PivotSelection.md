---
title: PivotTable.PivotSelection property (Excel)
keywords: vbaxl10.chm235124
f1_keywords:
- vbaxl10.chm235124
ms.prod: excel
api_name:
- Excel.PivotTable.PivotSelection
ms.assetid: efc3898f-aba8-3ffb-1421-da4c4864b712
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PivotSelection property (Excel)

Returns or sets the PivotTable selection in standard PivotTable report selection format. Read/write **String**.


## Syntax

_expression_.**PivotSelection**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Setting this property is equivalent to calling the **[PivotSelect](excel.pivottable.pivotselect.md)** method with the _Mode_ argument set to **xlDataAndLabel**.


## Example

This example selects the data and label for the salesperson named Bob in the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables(1).PivotSelection = "Salesman[Bob]"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]