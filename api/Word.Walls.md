---
title: Walls object (Word)
keywords: vbawd10.chm384
f1_keywords:
- vbawd10.chm384
ms.prod: word
api_name:
- Word.Walls
ms.assetid: e98c7218-b944-12bb-caf9-daecee4b6c0c
ms.date: 06/08/2017
localization_priority: Normal
---


# Walls object (Word)

Represents the walls of a 3D chart. 


## Remarks

This object is not a collection. There is no object that represents a single wall; you must return all the walls as a unit.


## Example

Use the **[Walls](Word.Chart.Walls.md)** property to return the **Walls** object. The following example sets the pattern on the walls for the first chart in the active document. If the chart is not a 3D chart, this example will fail.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Walls.Interior.Pattern = xlGray75 
 End If 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]