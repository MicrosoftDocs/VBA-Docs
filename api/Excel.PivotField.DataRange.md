---
title: PivotField.DataRange property (Excel)
keywords: vbaxl10.chm240078
f1_keywords:
- vbaxl10.chm240078
ms.prod: excel
api_name:
- Excel.PivotField.DataRange
ms.assetid: 14d5e4c4-1acb-aa02-6694-28e358afc881
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DataRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object as shown in the following table. Read-only.


## Syntax

_expression_.**DataRange**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

|Object|Data range|
|:-----|:-----|
|Data field|Data contained in the field|
|Row, column, or page field|Items in the field|
|Item|Data qualified by the item|

## Example

This example selects the PivotTable items in the field named REGION.

```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
Worksheets("Sheet1").Activate 
pvtTable.PivotFields("REGION").DataRange.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]