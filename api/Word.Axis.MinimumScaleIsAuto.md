---
title: Axis.MinimumScaleIsAuto property (Word)
keywords: vbawd10.chm113049634
f1_keywords:
- vbawd10.chm113049634
ms.prod: word
api_name:
- Word.Axis.MinimumScaleIsAuto
ms.assetid: 7e9ca498-1872-c4b1-e0b0-8d4444387747
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinimumScaleIsAuto property (Word)

 **True** if Microsoft Word calculates the minimum value for the value axis. Read/write **Boolean**.


## Syntax

_expression_. `MinimumScaleIsAuto`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Setting the  **[MinimumScale](Word.Axis.MinimumScale.md)** property sets this property to **False**.


## Example

The following example automatically calculates the minimum scale and the maximum scale for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MinimumScaleIsAuto = True 
 .MaximumScaleIsAuto = True 
 End With 
 End If 
End With 

```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]