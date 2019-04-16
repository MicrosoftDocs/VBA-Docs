---
title: Chart.Walls property (Word)
keywords: vbawd10.chm79364152
f1_keywords:
- vbawd10.chm79364152
ms.prod: word
api_name:
- Word.Chart.Walls
ms.assetid: f45ae75a-c96c-4441-af81-aedf23787194
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Walls property (Word)

Returns the walls of the 3D chart. Read-only  **[Walls](Word.Walls.md)**.


## Syntax

_expression_.**Walls**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example sets the color of the wall border of the first chart in the active document to red. You should run the example on a 3D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Walls.Border. _ 
 ColorIndex = 3 
 End If 
End With 

```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]