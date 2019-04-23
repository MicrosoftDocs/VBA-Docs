---
title: DataBar.BarBorder property (Excel)
keywords: vbaxl10.chm810094
f1_keywords:
- vbaxl10.chm810094
ms.prod: excel
api_name:
- Excel.DataBar.BarBorder
ms.assetid: d573e56e-cd02-c67e-ace8-8e8bdf2efd00
ms.date: 04/23/2019
localization_priority: Normal
---


# DataBar.BarBorder property (Excel)

Returns an object that specifies the border of a data bar. Read-only.


## Syntax

_expression_.**BarBorder**

_expression_ A variable that represents a **[DataBar](Excel.DataBar.md)** object.


## Return value

**[DataBarBorder](Excel.DataBarBorder.md)**


## Example

The following code example selects a range of cells, adds a data bar conditional formatting rule to that range, uses the **BarBorder** property to retrieve the **DataBarBorder** object associated with that rule, and then sets the data bar's color, tint, and type.

```vb
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
With myDataBar.BarBorder 
 .Type = xlDataBarBorderSolid 
 .Color.ThemeColor = xlThemeColorAccent2 
 .Color.TintAndShade = 0 
End With 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]