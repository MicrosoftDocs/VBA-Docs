---
title: Floor object (Word)
keywords: vbawd10.chm702
f1_keywords:
- vbawd10.chm702
ms.prod: word
api_name:
- Word.Floor
ms.assetid: 01d277eb-501b-09e5-65b8-83506c76ac05
ms.date: 06/08/2017
localization_priority: Normal
---


# Floor object (Word)

Represents the floor of a 3D chart.


## Example

Use the  **[Floor](Word.Chart.Floor.md)** property to return the **Floor** object. The following example sets the floor color for embedded chart one to cyan. The example will fail if the chart is not a 3D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Floor.Interior.Color = RGB(0, 255, 255) 
 End If 
End With 

```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]