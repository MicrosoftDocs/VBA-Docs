---
title: PivotTable.ErrorString property (Excel)
keywords: vbaxl10.chm235109
f1_keywords:
- vbaxl10.chm235109
ms.prod: excel
api_name:
- Excel.PivotTable.ErrorString
ms.assetid: 7f00d151-9f92-a3b3-c95f-60c0600cf594
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ErrorString property (Excel)

Returns or sets a **String** value that represents the string displayed in cells that contain errors when the **[DisplayErrorString](Excel.PivotTable.DisplayErrorString.md)** property is **True**.


## Syntax

_expression_.**ErrorString**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

The default value for this property is an empty string ("").


## Example

This example displays a hyphen in cells that contain errors in the specified PivotTable report.

```vb
With Worksheets(1).PivotTables("Pivot1") 
 .ErrorString = "-" 
 .DisplayErrorString = True 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]