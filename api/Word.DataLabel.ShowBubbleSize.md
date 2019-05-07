---
title: DataLabel.ShowBubbleSize property (Word)
keywords: vbawd10.chm233900010
f1_keywords:
- vbawd10.chm233900010
ms.prod: word
api_name:
- Word.DataLabel.ShowBubbleSize
ms.assetid: f3126ab6-7f58-d8f3-c0c4-6ace5e7dd8b7
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowBubbleSize property (Word)

 **True** to show the bubble size for the data labels on a chart. **False** to hide the bubble size. Read/write **Boolean**.


## Syntax

_expression_.**ShowBubbleSize**

_expression_ A variable that represents a '[DataLabel](Word.DataLabel.md)' object.


## Example

The following example shows the bubble size for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowBubbleSize = True 
 End If 
End With
```


## See also


[DataLabel Object](Word.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]