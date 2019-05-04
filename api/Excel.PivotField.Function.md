---
title: PivotField.Function property (Excel)
keywords: vbaxl10.chm240081
f1_keywords:
- vbaxl10.chm240081
ms.prod: excel
api_name:
- Excel.PivotField.Function
ms.assetid: 855334f6-dd6d-c09f-7732-c621751374a9
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.Function property (Excel)

Returns or sets the function used to summarize the PivotTable field (data fields only). Read/write **[XlConsolidationFunction](Excel.XlConsolidationFunction.md)**.


## Syntax

_expression_.**Function**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

For OLAP data sources, this property is read-only and always returns **xlUnknown**. For other data sources, this property cannot be set to **xlUnknown**.


## Example

This example sets the "Sum of 1994" field in the first PivotTable report on the active sheet to use the SUM function.

```vb
ActiveSheet.PivotTables("PivotTable1") _ 
 .PivotFields("Sum of 1994").Function = xlSum
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
