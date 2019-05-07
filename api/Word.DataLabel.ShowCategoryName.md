---
title: DataLabel.ShowCategoryName property (Word)
keywords: vbawd10.chm233900007
f1_keywords:
- vbawd10.chm233900007
ms.prod: word
api_name:
- Word.DataLabel.ShowCategoryName
ms.assetid: a2ef8f99-c26f-d0c1-4cd5-6a4787f69a0a
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabel.ShowCategoryName property (Word)

 **True** to display the category name for the data labels on a chart. **False** to hide the category name. Read/write **Boolean**.


## Syntax

_expression_.**ShowCategoryName**

_expression_ A variable that represents a '[DataLabel](Word.DataLabel.md)' object.


## Example

The following example shows the category name for the data labels of the first series on the first chart.


```vb
With ActiveDocument.InlineShapes(1) 
    If .HasChart Then 
        .Chart.SeriesCollection(1).DataLabels. _ 
            ShowCategoryName = True 
    End If 
End With
```


## See also


[DataLabel Object](Word.DataLabel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]