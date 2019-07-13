---
title: ColorFormat.ObjectThemeColor property (PowerPoint)
keywords: vbapp10.chm506006
f1_keywords:
- vbapp10.chm506006
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.ObjectThemeColor
ms.assetid: 40264b94-b16d-2a52-9adc-8e8510ec581d
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorFormat.ObjectThemeColor property (PowerPoint)

Returns or sets the theme color of the specified  **ColorFormat** object. Read/write.


## Syntax

_expression_. `ObjectThemeColor`

 _expression_ An expression that returns a [ColorFormat](PowerPoint.ColorFormat.md) object.


## Return value

MsoThemeColorIndex


## Remarks

The value of the  **ObjectThemeColor** property can be one of these **[MsoThemeColorIndex](Office.MsoThemeColorIndex.md)** constants.


## Example

The following example shows how to use the  **ObjectThemeColor** property to get the theme color of the foreground fill of shape one on slide one of the active presentation.


```vb
Public Sub ObjectThemeColor_Example() 
 
    Debug.Print ActivePresentation.Slides(1).Shapes(1).Fill.ForeColor.ObjectThemeColor 
     
End Sub
```


## See also


[ColorFormat Object](PowerPoint.ColorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]