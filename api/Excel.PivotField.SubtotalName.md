---
title: PivotField.SubtotalName property (Excel)
keywords: vbaxl10.chm240123
f1_keywords:
- vbaxl10.chm240123
ms.prod: excel
api_name:
- Excel.PivotField.SubtotalName
ms.assetid: db2f8366-75a4-edca-f46f-f0bff083ccbe
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.SubtotalName property (Excel)

Returns or sets the text string label displayed in the subtotal column or row heading in the specified PivotTable report. The default value is the string Subtotal. Read/write **String**.


## Syntax

_expression_.**SubtotalName**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Example

This example sets the subtotal label to Regional Subtotal (instead of the default string Subtotal) in the state field in the second PivotTable report on the active worksheet.

```vb
ActiveSheet.PivotTables("PivotTable2") _ 
 .PivotFields("state").SubtotalName = "Regional Subtotal"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]