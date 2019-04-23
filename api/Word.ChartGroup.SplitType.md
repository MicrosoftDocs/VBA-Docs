---
title: ChartGroup.SplitType property (Word)
keywords: vbawd10.chm263454760
f1_keywords:
- vbawd10.chm263454760
ms.prod: word
api_name:
- Word.ChartGroup.SplitType
ms.assetid: 0bebc2f8-4dd6-8a74-993b-9e16357f38d0
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.SplitType property (Word)

Returns or sets the way the two sections of either a pie-of-pie chart or a bar-of-pie chart are split. Read/write  **[XlChartSplitType](Word.xlchartsplittype.md)**.


## Syntax

_expression_.**SplitType**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Example

The following example splits the two sections of the chart by value, combining all values under 10 in the primary pie and displaying them in the secondary section. You must run the example on either a pie-of-pie chart or a bar-of-pie chart. 


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