---
title: FormatCondition.Borders property (Excel)
keywords: vbaxl10.chm512079
f1_keywords:
- vbaxl10.chm512079
api_name:
- Excel.FormatCondition.Borders
ms.assetid: 2f165a74-0b95-6643-5bd2-6a778523a411
ms.date: 10/12/2026
ms.localizationpriority: medium
---


# FormatCondition.Borders property (Excel)

Returns a **[Borders](Excel.Borders.md)** collection that represents the borders of a conditional format.


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a **[FormatCondition](Excel.FormatCondition.md)** object.


## Remarks

The **Borders** collection returned by a **FormatCondition** uses the conditional-format side constants **xlLeft**, **xlRight**, **xlTop**, and **xlBottom**.

Don't use **xlEdgeLeft**, **xlEdgeRight**, **xlEdgeTop**, **xlEdgeBottom**, **xlInsideHorizontal**, or **xlInsideVertical** with this collection. Those constants apply to **[Range.Borders](Excel.Range.Borders.md)**. Using an **xlEdge** index with **FormatCondition.Borders** can cause run-time error 1004 (COM error `0x800A03EC`) when you access properties such as **[LineStyle](Excel.Border.LineStyle.md)** or **[Weight](Excel.Border.Weight.md)**.

You can read or write the **LineStyle**, **Weight**, and color properties when you select a supported conditional-format side. For the supported line styles and weights, see the **[FormatCondition](Excel.FormatCondition.md)** object.


## Example

This example sets the bottom border of the first conditional-format rule for cell B2 on Sheet1 to a thin red border.

```vb
Sub SetConditionalFormatBorder()

    With Worksheets("Sheet1").Range("B2") _
            .FormatConditions(1).Borders(xlBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 3
    End With

End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]