---
title: DataBarBorder.Color property (Excel)
keywords: vbaxl10.chm885074
f1_keywords:
- vbaxl10.chm885074
ms.prod: excel
api_name:
- Excel.DataBarBorder.Color
ms.assetid: a16439a9-c086-9c42-8496-9a16d9011689
ms.date: 04/23/2019
localization_priority: Normal
---


# DataBarBorder.Color property (Excel)

Returns an object that specifies the color of the border of data bars specified by a conditional formatting rule. Read-only.


## Syntax

_expression_.**Color**

_expression_ A variable that represents a **[DataBarBorder](Excel.DataBarBorder.md)** object.


## Return value

**[FormatColor](Excel.FormatColor.md)**


## Example

The following code example selects a range of cells and adds a data bar conditional formatting rule to that range. It then uses the **[BarBorder](Excel.DataBar.BarBorder.md)** property to retrieve the **DataBarBorder** object associated with that rule, and uses the **Color** property of that object to retrieve the **FormatColor** object to set the color and tint of the data bar borders.

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