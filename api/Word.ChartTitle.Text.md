---
title: ChartTitle.Text property (Word)
keywords: vbawd10.chm65273868
f1_keywords:
- vbawd10.chm65273868
ms.prod: word
api_name:
- Word.ChartTitle.Text
ms.assetid: 4d17f47e-e2cb-fa62-fce1-27b70c7b8f70
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartTitle.Text property (Word)

Returns or sets the text for the specified object. Read/write  **String**.


## Syntax

_expression_.**Text**

_expression_ A variable that represents a '[ChartTitle](Word.ChartTitle.md)' object.


## Example

The following example sets the text for the chart title of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HasTitle = True 
 .Chart.ChartTitle.Text = "First Quarter Sales" 
 End If 
End With
```


## See also


[ChartTitle Object](Word.ChartTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]