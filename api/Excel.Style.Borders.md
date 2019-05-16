---
title: Style.Borders property (Excel)
keywords: vbaxl10.chm177075
f1_keywords:
- vbaxl10.chm177075
ms.prod: excel
api_name:
- Excel.Style.Borders
ms.assetid: 7da8309e-f01f-b131-b462-f974dde67007
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.Borders property (Excel)

Returns a **[Borders](Excel.Borders.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example sets the color of the bottom border of cell B2 on Sheet1 to a thin red border.

```vb
Sub SetRangeBorder() 
 
 With Worksheets("Sheet1").Range("B2").Borders(xlEdgeBottom) 
 .LineStyle = xlContinuous 
 .Weight = xlThin 
 .ColorIndex = 3 
 End With 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]