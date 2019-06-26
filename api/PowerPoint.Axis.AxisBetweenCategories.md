---
title: Axis.AxisBetweenCategories property (PowerPoint)
keywords: vbapp10.chm682001
f1_keywords:
- vbapp10.chm682001
ms.prod: powerpoint
api_name:
- PowerPoint.Axis.AxisBetweenCategories
ms.assetid: 8e0e0e80-58b9-005f-c719-ad45b491f9a9
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.AxisBetweenCategories property (PowerPoint)

 **True** if the value axis crosses the category axis between categories. Read/write **Boolean**.


## Syntax

_expression_.**AxisBetweenCategories**

_expression_ A variable that represents an '[Axis](PowerPoint.Axis.md)' object.


## Remarks

This property applies only to category axes, and it does not apply to 3D charts.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example causes the value axis for the first chart in the active document to cross the category axis between categories.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.Axes(xlCategory). _
            AxisBetweenCategories = True
    End If
End With
```


## See also


[Axis Object](PowerPoint.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]