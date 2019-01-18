---
title: AxisTitle.Text property (Word)
keywords: vbawd10.chm98238476
f1_keywords:
- vbawd10.chm98238476
ms.prod: word
api_name:
- Word.AxisTitle.Text
ms.assetid: 18aab6f0-84ec-0ec1-f1fd-82b0d6b114bd
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle.Text property (Word)

Returns or sets the text for the specified object. Read/write  **String**.


## Syntax

 _expression_. `Text`

 _expression_ A variable that represents an '[AxisTitle](Word.AxisTitle.md)' object.


## Example

The following example sets the axis title text for the category axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Text = "Month" 
 End With 
 End If 
End With
```


## See also


[AxisTitle Object](Word.AxisTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]