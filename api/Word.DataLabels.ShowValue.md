---
title: DataLabels.ShowValue property (Word)
keywords: vbawd10.chm207489000
f1_keywords:
- vbawd10.chm207489000
ms.prod: word
api_name:
- Word.DataLabels.ShowValue
ms.assetid: 3c016afc-17b2-78cd-8964-584e8d86d552
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabels.ShowValue property (Word)

 **True** to display the data label values for a specified chart. **False** to hide the values. Read/write **Boolean**.


## Syntax

_expression_.**ShowValue**

_expression_ A variable that represents a **[DataLabels](Word.DataLabels.md)** object.


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


[DataLabels Object](Word.DataLabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]