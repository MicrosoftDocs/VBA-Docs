---
title: ChartGroup.DoughnutHoleSize property (Word)
keywords: vbawd10.chm263454722
f1_keywords:
- vbawd10.chm263454722
ms.prod: word
api_name:
- Word.ChartGroup.DoughnutHoleSize
ms.assetid: 5f4098ee-7d94-ace4-b412-1c7071434973
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.DoughnutHoleSize property (Word)

Returns or sets the size of the hole in a doughnut chart group. Read/write  **Long**.


## Syntax

_expression_.**DoughnutHoleSize**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

The hole size is expressed as a percentage of the chart size, from 10 through 90 percent.


## Example

The following example sets the hole size for doughnut group one of the first chart in the active document. You should run the example on a 2D doughnut chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.DoughnutGroups(1).DoughnutHoleSize = 10 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]