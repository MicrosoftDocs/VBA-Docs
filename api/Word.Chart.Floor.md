---
title: Chart.Floor property (Word)
keywords: vbawd10.chm79364153
f1_keywords:
- vbawd10.chm79364153
ms.prod: word
api_name:
- Word.Chart.Floor
ms.assetid: 1544e584-3b0f-92a8-cc8f-6b12ed66109e
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Floor property (Word)

Returns the floor of the 3D chart. Read-only  **[Floor](Word.Floor.md)**.


## Syntax

_expression_.**Floor**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example sets the floor color of the first chart in the active document to blue. You should run the example on a 3D chart (the  **Floor** property fails on 2D charts).


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Floor.Interior.ColorIndex = 5 
 End If 
End With 

```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]