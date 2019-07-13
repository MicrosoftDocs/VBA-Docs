---
title: Slide.Background property (PowerPoint)
keywords: vbapp10.chm531007
f1_keywords:
- vbapp10.chm531007
ms.prod: powerpoint
api_name:
- PowerPoint.Slide.Background
ms.assetid: 8af622b9-029a-6839-7a44-fdf96fe75dc9
ms.date: 06/08/2017
localization_priority: Normal
---


# Slide.Background property (PowerPoint)

Returns a  **[ShapeRange](PowerPoint.ShapeRange.md)** object that represents the slide background.


## Syntax

_expression_.**Background**

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Return value

ShapeRange


## Remarks

If you use the  **Background** property to set the background for an individual slide without changing the slide master, the **FollowMasterBackground** property for that slide must be set to **False**.


## Example

This example sets the background of the slide master in the active presentation to a preset shade.


```vb
ActivePresentation.SlideMaster.Background.Fill.PresetGradient _
    Style:=msoGradientHorizontal, Variant:=1, _
    PresetGradientType:=msoGradientLateSunset
```

This example sets the background of slide one in the active presentation to a preset shade.




```vb
With ActivePresentation.Slides(1)
    .FollowMasterBackground = False
    .Background.Fill.PresetGradient Style:=msoGradientHorizontal, _
        Variant:=1, PresetGradientType:=msoGradientLateSunset
End With
```


## See also


[Slide Object](PowerPoint.Slide.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]