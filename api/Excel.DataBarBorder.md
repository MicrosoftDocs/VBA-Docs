---
title: DataBarBorder object (Excel)
keywords: vbaxl10.chm884072
f1_keywords:
- vbaxl10.chm884072
ms.prod: excel
api_name:
- Excel.DataBarBorder
ms.assetid: e46bb88b-ec41-a4f9-8926-34d0a22ad8e9
ms.date: 03/29/2019
localization_priority: Normal
---


# DataBarBorder object (Excel)

Represents the border of the data bars specified by a conditional formatting rule.


## Remarks

Use the **DataBarBorder** object to get or set the color and border type for data bars. To access the **DataBarBorder** object associated with a data bar conditional formatting rule, use the **[BarBorder](Excel.DataBar.BarBorder.md)** property. After retrieving the **DataBarBorder** object, use its **Color** property to return a **[FormatColor](Excel.FormatColor.md)** object that you can use to set the color of the data bars.


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


## Properties

- [Application](Excel.DataBarBorder.Application.md)
- [Color](Excel.DataBarBorder.Color.md)
- [Creator](Excel.DataBarBorder.Creator.md)
- [Parent](Excel.DataBarBorder.Parent.md)
- [Type](Excel.DataBarBorder.Type.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]