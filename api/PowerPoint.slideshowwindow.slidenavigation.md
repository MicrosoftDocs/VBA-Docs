---
title: SlideShowWindow.SlideNavigation property (PowerPoint)
keywords: vbapp10.chm507013
f1_keywords:
- vbapp10.chm507013
ms.assetid: 232fa845-0486-5288-fd27-dc41d83096e1
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# SlideShowWindow.SlideNavigation property (PowerPoint)

Returns a  **[SlideNavigation](PowerPoint.slidenavigation.md)** object that represents the slide navigation screen in slide show view. Read-only


## Syntax

_expression_. `SlideNavigation`

_expression_ A variable that represents a [SlideShowWindow](PowerPoint.SlideShowWindow.md) object.


## Example

The following code sample starts a slide show from the active presentation and then makes the navigation screen visible.


```vb
Sub ShowSlideNavigation()

    ' Start the slide show.
    ActivePresentation.SlideShowSettings.Run
    
    ' Show the slide navigation screen.
    ActivePresentation.SlideShowWindow.SlideNavigation.Visible = True

End Sub
```


## Property value

 **SLIDENAVIGATION**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]