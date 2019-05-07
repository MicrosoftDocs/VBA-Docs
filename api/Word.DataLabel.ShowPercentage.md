---
title: DataLabel.ShowPercentage property (Word)
keywords: vbawd10.chm233900009
f1_keywords:
- vbawd10.chm233900009
ms.prod: word
api_name:
- Word.DataLabel.ShowPercentage
ms.assetid: 4347e76f-0107-f153-ab4b-5897683d6495
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowPercentage property (Word)

 **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean**.


## Syntax

_expression_.**ShowPercentage**

_expression_ A variable that represents a '[DataLabel](Word.DataLabel.md)' object.


## Example

The following example enables the percentage value to be shown for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowPercentage = True 
 End If 
End With
```


## See also


[DataLabel Object](Word.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]