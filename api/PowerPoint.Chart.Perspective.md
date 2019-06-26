---
title: Chart.Perspective property (PowerPoint)
keywords: vbapp10.chm684037
f1_keywords:
- vbapp10.chm684037
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Perspective
ms.assetid: 0ac63aba-4182-c8dc-d51b-a75539025865
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Perspective property (PowerPoint)

Returns or sets the perspective for the 3D chart view. Read/write  **Long**.


## Syntax

_expression_.**Perspective**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

The value of this property must be between 0 and 100. This property is ignored if the  **[RightAngleAxes](PowerPoint.Chart.RightAngleAxes.md)** property is set to **True**.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

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


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]