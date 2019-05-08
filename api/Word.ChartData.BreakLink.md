---
title: ChartData.BreakLink method (Word)
keywords: vbawd10.chm190382083
f1_keywords:
- vbawd10.chm190382083
ms.prod: word
api_name:
- Word.ChartData.BreakLink
ms.assetid: 19b483c2-8fca-38f5-c769-f7052c3bfee1
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartData.BreakLink method (Word)

Removes the link between the data for a chart and a Microsoft Excel workbook.


## Syntax

_expression_. `BreakLink`

_expression_ A variable that represents a '[ChartData](Word.ChartData.md)' object.


## Remarks

Calling this method sets the  **[IsLinked](Word.ChartData.IsLinked.md)** property of the **ChartData** object to **False**.


## Example

The following example removes the link between the  **ChartData** object for the first chart in the active document and the Excel workbook that provided the data for the chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartData.Activate 
 .Chart.ChartData.BreakLink 
 End If 
End With
```


## See also


[ChartData Object](Word.ChartData.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]