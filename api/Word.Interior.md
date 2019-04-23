---
title: Interior object (Word)
keywords: vbawd10.chm43
f1_keywords:
- vbawd10.chm43
ms.prod: word
api_name:
- Word.Interior
ms.assetid: 6fc3e311-a7c9-bfa9-7459-9cea177b08e5
ms.date: 06/08/2017
localization_priority: Normal
---


# Interior object (Word)

Represents the interior of an object.


## Example

The following example enables up and down bars, and then sets the interior color of the up bars to green, for the first chart group of the first chart in the active document. Use the  **[UpBars.Interior](overview/Word.md)** property to return the **Interior** object.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .UpBars.Interior.ColorIndex = 4 
 End With 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]