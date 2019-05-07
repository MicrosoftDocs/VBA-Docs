---
title: PivotField.TotalLevels property (Excel)
keywords: vbaxl10.chm240097
f1_keywords:
- vbaxl10.chm240097
ms.prod: excel
api_name:
- Excel.PivotField.TotalLevels
ms.assetid: fa50c186-5f6d-41f4-6382-37135159347c
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotField.TotalLevels property (Excel)

Returns the total number of fields in the current field group. If the field isn't grouped, or if the data source is OLAP-based, **TotalLevels** returns the value 1. Read-only **Long**.


## Syntax

_expression_.**TotalLevels**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

All fields in a set of grouped fields have the same **TotalLevels** value.


## Example

This example displays the total number of fields in the group that contains the active cell.

```vb
Worksheets("Sheet1").Activate 
MsgBox "This group has " & _ 
 ActiveCell.PivotField.TotalLevels & " levels
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]