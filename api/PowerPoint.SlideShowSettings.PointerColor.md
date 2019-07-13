---
title: SlideShowSettings.PointerColor property (PowerPoint)
keywords: vbapp10.chm514003
f1_keywords:
- vbapp10.chm514003
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.PointerColor
ms.assetid: 530072d6-3a2d-8236-b4ac-3ede8823e95a
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.PointerColor property (PowerPoint)

Returns the pointer color for the specified presentation as a  **[ColorFormat](PowerPoint.ColorFormat.md)** object. Read-only.


## Syntax

_expression_. `PointerColor`

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

ColorFormat


## Remarks

The pointer color is saved with the presentation and is the default pen color each time you show the presentation. 

To change the pointer to a pen, set the  **[PointerType](PowerPoint.SlideShowView.PointerType.md)** property to **ppSlideShowPointerPen**.


## Example

This example sets the default pen color for the active presentation to blue, starts a slide show, changes the pointer to a pen, and then sets the pen color to red for this slide show only.


```vb
With ActivePresentation.SlideShowSettings

    .PointerColor.RGB = RGB(0, 0, 255)          'blue

    With .Run.View

        .PointerColor.RGB = RGB(255, 0, 0)      'red

        .PointerType = ppSlideShowPointerPen

    End With

End With
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]