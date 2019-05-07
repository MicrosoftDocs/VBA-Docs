---
title: DataLabel.ShowValue property (Word)
keywords: vbawd10.chm233900008
f1_keywords:
- vbawd10.chm233900008
ms.prod: word
api_name:
- Word.DataLabel.ShowValue
ms.assetid: 1dec8c2c-07b0-57a1-7f66-da0d263d6075
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowValue property (Word)

 **True** to display a specified chart's data label values. **False** to hide the values. Read/write **Boolean**.


## Syntax

_expression_.**ShowValue**

_expression_ A variable that represents a '[DataLabel](Word.DataLabel.md)' object.


## Example

The following example enables the value to be shown for the data labels of the first series in the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowValue = True 
 End If 
End With
```


## See also


[DataLabel Object](Word.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]