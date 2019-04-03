---
title: ChartTitle object (Word)
keywords: vbawd10.chm996
f1_keywords:
- vbawd10.chm996
ms.prod: word
api_name:
- Word.ChartTitle
ms.assetid: fc8ca540-0a29-123b-2fdf-b16aaa1f940c
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle object (Word)

Represents the chart title.


## Remarks

Use the  **[ChartTitle](Word.Chart.ChartTitle.md)** property to return the **ChartTitle** object.

The  **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](Word.Chart.HasTitle.md)** property for the chart is **True**.


## Example

 The following example adds a title to embedded chart one on the worksheet named **Sheet1**.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "February Sales" 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]