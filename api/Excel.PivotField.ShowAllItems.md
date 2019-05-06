---
title: PivotField.ShowAllItems property (Excel)
keywords: vbaxl10.chm240088
f1_keywords:
- vbaxl10.chm240088
ms.prod: excel
api_name:
- Excel.PivotField.ShowAllItems
ms.assetid: 8dc34e02-bdfb-6972-04fa-22ba1977c0c8
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.ShowAllItems property (Excel)

**True** if all items in the PivotTable report are displayed, even if they don't contain summary data. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**ShowAllItems**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

For OLAP data sources, the value is always **False**.


## Example

This example displays all rows for the Month field in the first PivotTable report on worksheet one, including months for which there's no data.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Month").ShowAllItems = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
