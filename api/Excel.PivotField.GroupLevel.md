---
title: PivotField.GroupLevel property (Excel)
keywords: vbaxl10.chm240082
f1_keywords:
- vbaxl10.chm240082
ms.prod: excel
api_name:
- Excel.PivotField.GroupLevel
ms.assetid: fc017652-bded-4655-03df-79cfa733b12e
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.GroupLevel property (Excel)

Returns the placement of the specified field within a group of fields (if the field is a member of a grouped set of fields). Read-only.


## Syntax

_expression_.**GroupLevel**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

This property is not available for OLAP data sources.

The highest-level parent field (leftmost parent field) is level one, its child is level two, and so on.


## Example

This example displays a message box if the field that contains the active cell is the highest-level parent field.

```vb
Worksheets("Sheet1").Activate 
If ActiveCell.PivotField.GroupLevel = 1 Then 
 MsgBox "This is the highest-level parent field." 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]