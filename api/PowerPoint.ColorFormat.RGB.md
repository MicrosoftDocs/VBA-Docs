---
title: ColorFormat.RGB property (PowerPoint)
keywords: vbapp10.chm506002
f1_keywords:
- vbapp10.chm506002
api_name:
- PowerPoint.ColorFormat.RGB
ms.assetid: 5bb68052-5931-2096-277c-fb44c76b37eb
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# ColorFormat.RGB property (PowerPoint)

Returns or sets the red-green-blue (RGB) value of the specified color. Read/write.


## Syntax

_expression_. `RGB`

_expression_ A variable that represents a [ColorFormat](PowerPoint.ColorFormat.md) object.


## Return value

Long


## Example

This example sets the background color for the first shape on the first slide of the active presentation to red and Accent 1 to green in the theme for the first slide master.


```vb
With ActivePresentation
    Dim oCF As ColorFormat
    
    Set oCF = .Slides(1).Shapes(1).Fill.ForeColor
    
    oCF.RGB = RGB(255, 0, 0)

    .Designs(1).SlideMaster.Theme.ThemeColorScheme(msoThemeAccent1).RGB = RGB(0, 255, 0)
End With
```


## See also


[ColorFormat Object](PowerPoint.ColorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
