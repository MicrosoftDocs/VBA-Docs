---
title: ChartGroup.SplitValue property (Word)
keywords: vbawd10.chm263454762
f1_keywords:
- vbawd10.chm263454762
ms.prod: word
api_name:
- Word.ChartGroup.SplitValue
ms.assetid: 102826a5-834e-1b23-9888-6fb9b193ac96
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SplitValue property (Word)

Returns or sets the threshold value separating the two sections of either a pie-of-pie chart or a bar-of-pie chart. Read/write  **Variant**.


## Syntax

_expression_.**SplitValue**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Example

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. You must run this example on either a pie-of-pie chart or a bar-of-pie chart. 


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .SplitType = xlSplitByValue 
 .SplitValue = 10 
 .VaryByCategories = True 
 End With 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]