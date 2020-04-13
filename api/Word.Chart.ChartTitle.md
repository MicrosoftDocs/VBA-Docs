---
title: Chart.ChartTitle property (Word)
keywords: vbawd10.chm79364099
f1_keywords:
- vbawd10.chm79364099
ms.prod: word
api_name:
- Word.Chart.ChartTitle
ms.assetid: 1804d06a-bb2b-5995-7750-2ada70ddd1d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartTitle property (Word)

Returns the title of the specified chart. Read-only  **[ChartTitle](Word.ChartTitle.md)**.


## Syntax

_expression_. `ChartTitle`

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Remarks

The **ChartTitle** object does not exist and cannot be used unless the **[HasTitle](Word.Chart.HasTitle.md)** property for the chart is **True**.


## Example

The following example sets the text for the title of the first chart.


```vb
With ActiveDocument.InlineShapes(1).Chart 
 .HasTitle = True 
 .ChartTitle.Text = "First Quarter Sales" 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]