---
title: Chart.Elevation property (PowerPoint)
keywords: vbapp10.chm684027
f1_keywords:
- vbapp10.chm684027
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.Elevation
ms.assetid: 9b6ac583-2a35-8737-5660-51d2b7e6dbbd
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Elevation property (PowerPoint)

Returns or sets the elevation, in degrees, of the 3D chart view. Read/write  **Long**.


## Syntax

_expression_.**Elevation**

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Remarks

The chart elevation is the height, in degrees, at which you view the chart. The default is 15 for most chart types. The value of this property must be between -90 and 90, except for 3D bar charts, where it must be between 0 and 44.


> [!NOTE] 
> This property applies only to 3D charts.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the chart elevation of the first chart in the active document to 34 degrees. You should run the example on a 3D chart.




```vb
With ActiveDocument.InlineShapes(1)

    If .HasChart Then

        .Chart.Elevation = 34

    End If

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]