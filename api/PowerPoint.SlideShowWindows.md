---
title: SlideShowWindows object (PowerPoint)
keywords: vbapp10.chm510000
f1_keywords:
- vbapp10.chm510000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowWindows
ms.assetid: aa4c7a38-32ea-c206-ce1f-d78094410f52
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowWindows object (PowerPoint)

A collection of all the  **[SlideShowWindow](PowerPoint.SlideShowWindow.md)** objects that represent the open slide shows in Microsoft PowerPoint.


## Example

Use the [SlideShowWindows](PowerPoint.Application.SlideShowWindows.md)property to return the  **SlideShowWindows** collection. Use **SlideShowWindows** (_index_), where _index_ is the window index number, to return a single **SlideShowWindow** object. The following example reduces the height of slide show window one if it is a full-screen window.


```vb
With SlideShowWindows(1)

    If .IsFullScreen Then

        .Height = .Height - 20

    End If

End With
```

Use the [Run](PowerPoint.SlideShowSettings.Run.md)method to create a new slide show window and add it to the  **SlideShowWindows** collection. The following example runs a slide show of the active presentation.




```vb
ActivePresentation.SlideShowSettings.Run
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]