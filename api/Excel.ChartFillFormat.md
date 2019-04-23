---
title: ChartFillFormat object (Excel Graph)
keywords: vbagr10.chm5207187
f1_keywords:
- vbagr10.chm5207187
ms.prod: excel
api_name:
- Excel.ChartFillFormat
ms.assetid: e011f58f-141b-1b21-0db4-04a5c5e964c6
ms.date: 04/06/2019
localization_priority: Normal
---


# ChartFillFormat object (Excel Graph)

Represents fill formatting.


## Remarks

Use the **[Fill](Excel.Fill.md)** property to return the **ChartFillFormat** object. 


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