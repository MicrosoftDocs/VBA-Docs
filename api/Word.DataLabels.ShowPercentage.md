---
title: DataLabels.ShowPercentage property (Word)
keywords: vbawd10.chm207489001
f1_keywords:
- vbawd10.chm207489001
ms.prod: word
api_name:
- Word.DataLabels.ShowPercentage
ms.assetid: d13c6988-d751-e084-8fc0-830cc1382906
ms.date: 06/08/2017
localization_priority: Normal
---


# DataLabels.ShowPercentage property (Word)

 **True** to display the percentage value for the data labels on a chart. **False** to hide the value. Read/write **Boolean**.


## Syntax

_expression_.**ShowPercentage**

_expression_ A variable that represents a **[DataLabels](Word.DataLabels.md)** object.


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


[DataLabels Object](Word.DataLabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]