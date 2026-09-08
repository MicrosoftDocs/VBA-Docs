---
title: FormatCondition.Borders property (Excel)
keywords: vbaxl10.chm512079
f1_keywords:
- vbaxl10.chm512079
api_name:
- Excel.FormatCondition.Borders
ms.assetid: 2f165a74-0b95-6643-5bd2-6a778523a411
ms.date: 09/03/2026
ms.localizationpriority: medium
---


# FormatCondition.Borders property (Excel)

Returns a **[Borders](Excel.Borders.md)** collection that represents the borders of a style or a range of cells (including a range defined as part of a conditional format).


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a **[FormatCondition](Excel.FormatCondition.md)** object.


## Remarks

For **Border** objects returned by this collection, you can set supported border properties for the top, bottom, left, and right borders. However, reading the **[LineStyle](Excel.Border.LineStyle.md)** or **[Weight](Excel.Border.Weight.md)** property isn't supported and can raise run-time error 1004 (COM error `0x800A03EC`), even when the border was configured in the Conditional Formatting dialog.

This limitation doesn't apply to **[Range.Borders](Excel.Range.Borders.md)**.


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

This example sets the bottom border of the first conditional-format rule for cell B2 on Sheet1 to a thin red border.

```vb
Sub SetConditionalFormatBorder()

 With Worksheets("Sheet1").Range("B2").FormatConditions(1).Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = 3
 End With

End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]