---
title: ChartFont object (Word)
keywords: vbawd10.chm3905
f1_keywords:
- vbawd10.chm3905
ms.prod: word
api_name:
- Word.ChartFont
ms.assetid: 2ca7fb97-fa22-dec1-6978-8ebb6d8aad7c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartFont object (Word)

Contains the font attributes (font name, font size, color, and so on) for an object chart.


## Remarks

If you do not want to format all the text in an  **[AxisTitle](Word.AxisTitle.md)**, **[ChartTitle](Word.ChartTitle.md)**, **[DataLabel](Word.DataLabel.md)**, or **[DisplayUnitLabel](Word.DisplayUnitLabel.md)** object the same way, use the **Characters** property of that object to first return a subset of the text as a **[ChartCharacters](Word.ChartCharacters.md)** object. Then use the **[Font](Word.ChartCharacters.Font.md)** property of the **ChartCharacters** object to return a **ChartFont** object you can use to format the subset of text, as needed.


## Example

The following example formats the title of the first chart as bold. Use the **Font** property to return the **ChartFont** object.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .AxisTitle.Font.Bold = True 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]