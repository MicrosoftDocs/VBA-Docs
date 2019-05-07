---
title: PivotTable.EnableFieldDialog property (Excel)
keywords: vbaxl10.chm235107
f1_keywords:
- vbaxl10.chm235107
ms.prod: excel
api_name:
- Excel.PivotTable.EnableFieldDialog
ms.assetid: 4b6b4bc5-9b87-efa2-c6d1-4ab0c11f5966
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.EnableFieldDialog property (Excel)

**True** if the **PivotTable Field** dialog box is available when the user double-clicks the PivotTable field. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableFieldDialog**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Setting this property for a PivotTable report sets it for all fields in that report.


## Example

This example disables the **PivotTable Field** dialog box for the **Year** field.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").EnableFieldDialog = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]