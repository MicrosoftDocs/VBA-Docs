---
title: PivotField.AutoSort method (Excel)
keywords: vbaxl10.chm240157
f1_keywords:
- vbaxl10.chm240157
ms.prod: excel
api_name:
- Excel.PivotField.AutoSort
ms.assetid: 7a0bba4d-b18c-04df-a3b4-6ae2807f5238
ms.date: 06/08/2017
localization_priority: Priority
---


# PivotField.AutoSort method (Excel)

Establishes automatic field-sorting rules for PivotTable reports.


## Syntax

_expression_. `AutoSort`( `_Order_` , `_Field_` , `_PivotLine_` , `_CustomSubtotal_` )

_expression_ A variable that represents a [PivotField](Excel.PivotField.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Order_|Required| **Long**|One of the constants of  **[xlSortOrder](Excel.XlSortOrder.md)** specifying the sort order.|
| _Field_|Required| **String**|The name of the sort key field. You must specify the unique name (as returned from the  **[SourceName](Excel.PivotField.SourceName.md)** property), and not the displayed name.|
| _PivotLine_|Optional| **Variant**|A line on a column or row in a PivotTable report.|
| _CustomSubtotal_|Optional| **Variant**|The custom subtotal field.|

## Example

This example sorts the Company field in descending order, based on the sum of sales.


```vb
ActiveSheet.PivotTables(1).PivotField("Company") _ 
 .AutoSort xlDescending, "Sum of Sales"
```


## See also


[PivotField Object](Excel.PivotField.md)

