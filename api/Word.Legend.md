---
title: Legend object (Word)
keywords: vbawd10.chm2246
f1_keywords:
- vbawd10.chm2246
ms.prod: word
api_name:
- Word.Legend
ms.assetid: f0122074-87b7-0225-3c6c-406103fa4c29
ms.date: 06/08/2017
localization_priority: Normal
---


# Legend object (Word)

Represents the legend in a chart. Each chart can have only one legend.


## Remarks

 The **Legend** object contains one or more **[LegendEntry](Word.LegendEntry.md)** objects; each **LegendEntry** object contains a **[LegendKey](Word.LegendKey.md)** object.

The chart legend is not visible unless the **[HasLegend](Word.Chart.HasLegend.md)** property is **True**. If this property is **False**, properties and methods of the **Legend** object will fail.


## Example

Use the **[Legend](Word.Chart.Legend.md)** property to return the **Legend** object. The following example sets the font style for the legend of the first chart in the active document to bold.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Legend.Font.Bold = True 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]