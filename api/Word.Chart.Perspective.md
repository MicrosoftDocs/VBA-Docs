---
title: Chart.Perspective property (Word)
keywords: vbawd10.chm79364108
f1_keywords:
- vbawd10.chm79364108
ms.prod: word
api_name:
- Word.Chart.Perspective
ms.assetid: d88ab2dc-822a-de51-a2b9-bcce667cd0e2
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Perspective property (Word)

Returns or sets the perspective for the 3D chart view. Read/write  **Long**.


## Syntax

_expression_.**Perspective**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Remarks

The value of this property must be between 0 and 100. This property is ignored if the  **[RightAngleAxes](Word.Chart.RightAngleAxes.md)** property is set to **True**.


## Example

The following example sets the perspective of the first chart in the active document to 70. You should run the example on a 3D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.RightAngleAxes = False 
 .Chart.Perspective = 70 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]