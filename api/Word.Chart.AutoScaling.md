---
title: Chart.AutoScaling property (Word)
keywords: vbawd10.chm79364159
f1_keywords:
- vbawd10.chm79364159
ms.prod: word
api_name:
- Word.Chart.AutoScaling
ms.assetid: 911bf146-f3fa-7c05-a0eb-9e2062ed4a93
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.AutoScaling property (Word)

 **True** if Microsoft Word scales a 3D chart so that it is closer in size to the equivalent 2D chart. The **[RightAngleAxes](Word.Chart.RightAngleAxes.md)** property must be **True**. Read/write **Boolean**.


## Syntax

_expression_. `AutoScaling`

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example automatically scales the first chart in the active document. The example should be run on a 3D chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.RightAngleAxes = True 
 .Chart.AutoScaling = True 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]