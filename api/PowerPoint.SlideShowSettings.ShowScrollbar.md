---
title: SlideShowSettings.ShowScrollbar property (PowerPoint)
keywords: vbapp10.chm514015
f1_keywords:
- vbapp10.chm514015
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.ShowScrollbar
ms.assetid: 9f6be3f3-1099-2f8c-4c1c-b5ab1be89f4a
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.ShowScrollbar property (PowerPoint)

Determines whether to display the scroll bar during a slide show in browse mode. Read/write.


## Syntax

_expression_. `ShowScrollbar`

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

MsoTriState


## Remarks

Use the  **[ShowType](PowerPoint.SlideShowSettings.ShowType.md)** property prior to setting the **ShowScrollbar** property.

The value of the  **ShowScrollbar** property can be one of these **MsoTriState** constants.


||
|:-----|
|**msoFalse**|
|**msoTrue**|

## Example

This example specifies to display the slide show for the active presentation in a window and displays a scrollbar used for browsing through the slides during a slide show.


```vb
Sub ShowSlideShowScrollBar()

    With ActivePresentation.SlideShowSettings

        .ShowType = ppShowTypeWindow

        .ShowScrollBar = msoTrue

    End With

End Sub
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]