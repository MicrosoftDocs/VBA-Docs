---
title: ChartGroup.VaryByCategories property (Word)
keywords: vbawd10.chm263454752
f1_keywords:
- vbawd10.chm263454752
ms.prod: word
api_name:
- Word.ChartGroup.VaryByCategories
ms.assetid: e7ee35a4-ddb7-83ef-3c9b-0076f601bb19
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.VaryByCategories property (Word)

 **True** if Microsoft Word assigns a different color or pattern to each data marker. Read/write **Boolean**.


## Syntax

_expression_.**VaryByCategories**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

The chart must contain only one series. 


## Example

The following example assigns a different color or pattern to each data marker in chart group one. You should run the example on a 2D line chart that has data markers on a series.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).VaryByCategories = True 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]