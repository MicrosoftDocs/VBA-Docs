---
title: CubeField.PivotFields property (Excel)
keywords: vbaxl10.chm668088
f1_keywords:
- vbaxl10.chm668088
ms.prod: excel
api_name:
- Excel.CubeField.PivotFields
ms.assetid: d3da6064-a4b2-7075-cc3e-033896f5b4a9
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.PivotFields property (Excel)

Returns the **[PivotFields](Excel.PivotFields.md)** collection. This collection contains all PivotTable fields, including those that aren't currently visible on-screen. Read-only **PivotFields** object.


## Syntax

_expression_.**PivotFields**

_expression_ An expression that returns a **[CubeField](Excel.CubeField.md)** object.


## Return value

PivotFields


## Remarks

For Online Analytical Processing (OLAP) data sources, there are no hidden fields, and the object or collection that's returned reflects what's currently visible.


## Example

This example creates a list of all the PivotTable field names used in the first PivotChart report.

```vb
Set objNewSheet = Worksheets.Add 
objNewSheet.Activate 
intRow = 1 
For Each objPF In _ 
 Charts("Chart1").PivotLayout.PivotFields 
 objNewSheet.Cells(intRow, 1).Value = objPF.Caption 
 intRow = intRow + 1 
Next objPF
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]