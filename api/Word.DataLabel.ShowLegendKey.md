---
title: DataLabel.ShowLegendKey property (Word)
keywords: vbawd10.chm233898155
f1_keywords:
- vbawd10.chm233898155
ms.prod: word
api_name:
- Word.DataLabel.ShowLegendKey
ms.assetid: b9238117-ad3f-7dd7-bf35-d773bf713535
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowLegendKey property (Word)

 **True** if the data label legend key is visible. Read/write **Boolean**.


## Syntax

_expression_.**ShowLegendKey**

_expression_ A variable that represents a '[DataLabel](Word.DataLabel.md)' object.


## Example

The following example sets the data labels for series one of the first chart in the active document to show values and the legend key.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).DataLabels. _ 
 ShowLegendKey = True 
 .Chart.SeriesCollection(1).DataLabels.Type = xlShowValue 
 End If 
End With
```


## See also


[DataLabel Object](Word.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]