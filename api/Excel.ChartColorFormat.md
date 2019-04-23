---
title: ChartColorFormat object (Excel Graph)
keywords: vbagr10.chm131251
f1_keywords:
- vbagr10.chm131251
ms.prod: excel
api_name:
- Excel.ChartColorFormat
ms.assetid: 5d2e0cb0-e928-0704-7b4c-1afee6096f3a
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartColorFormat object (Excel Graph)

Represents a foreground or background color.


## Remarks

Use the **[ForeColor](Excel.ForeColor.md)** property to return a **ChartColorFormat** object that represents the foreground fill color. 

Use the **[BackColor](Excel.backcolor.md)** property to return the background fill color. 

Use the **[RGB](Excel.RGB.md)** property to return the color as an explicit red-green-blue value.

Use the **[SchemeColor](Excel.SchemeColor.md)** property to return or set the color as one of the colors in the current color scheme. 

## Example

The following example sets the foreground color, background color, and gradient for the chart area fill in _myChart_.

```vb
With myChart.ChartArea.Fill 
    .Visible = True 
    .ForeColor.SchemeColor = 15 
    .BackColor.SchemeColor = 17 
    .TwoColorGradient msoGradientHorizontal, 1 
End With
```

## See also

- [Excel Graph Visual Basic Reference](overview/excel/graph-visual-basic-reference.md)
- [Excel Object Model Reference](overview/excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]