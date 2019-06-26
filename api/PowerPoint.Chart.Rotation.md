---
title: Chart.Rotation property (PowerPoint)
keywords: vbapp10.chm684041
f1_keywords:
- vbapp10.chm684041
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Rotation
ms.assetid: 5f533c86-369c-6dbd-f70c-c7de0cc6d868
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Rotation property (PowerPoint)

Returns or sets the rotation, in degrees, of the 3D chart view (the rotation of the plot area around the z-axis). Read/write  **Variant**.


## Syntax

_expression_.**Rotation**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

The value of this property must be from 0 through 360, except for 3D bar charts, where the value must be from 0 through 44. The default value is 20. This property applies only to 3D charts. 

Rotations are always rounded to the nearest integer.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the rotation of the first chart in the active document to 30 degrees. You should run the example on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Rotation = 30

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]