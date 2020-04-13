---
title: ChartData object (Word)
keywords: vbawd10.chm2905
f1_keywords:
- vbawd10.chm2905
ms.prod: word
api_name:
- Word.ChartData
ms.assetid: 323ee62c-9b70-8280-d448-79cf4d2b6953
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData object (Word)

Represents access to the linked or embedded data associated with a chart.


## Remarks

Use the **[ChartData](Word.Chart.ChartData.md)** property to return the **ChartData** object.


## Example

The following example uses the **[Activate](Word.ChartData.Activate.md)** method to display the data associated with the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData 
 .Activate 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]