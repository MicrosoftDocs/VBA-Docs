---
title: SlideShowSettings.Run method (PowerPoint)
keywords: vbapp10.chm514008
f1_keywords:
- vbapp10.chm514008
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.Run
ms.assetid: 497fae3b-b6a3-dc26-20d9-bdc8057ddc09
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.Run method (PowerPoint)

Runs a slide show of the specified presentation. Returns a **[SlideShowWindow](PowerPoint.SlideShowWindow.md)** object.


## Syntax

_expression_.**Run**

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

SlideShowWindow


## Remarks

To run a custom slide show, set the  **RangeType** property to **ppShowNamedSlideShow**, and set the **SlideShowName** property to the name of the custom show you want to run.


## Example

This example starts a full-screen slide show of the active presentation, with shortcut keys disabled.


```vb
With ActivePresentation.SlideShowSettings

    .ShowType = ppShowSpeaker

    .Run.View.AcceleratorsEnabled = False

End With
```

This example runs the named slide show "Quick Show."




```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowNamedSlideShow

    .SlideShowName = "Quick Show"

    .Run

End With
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
