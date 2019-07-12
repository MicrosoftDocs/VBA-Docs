---
title: SlideShowView object (PowerPoint)
keywords: vbapp10.chm513000
f1_keywords:
- vbapp10.chm513000
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView
ms.assetid: 403b30ef-b12f-3a3c-e8d8-19189fd762fe
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView object (PowerPoint)

Represents the view in a slide show window.


## Example

Use the [View](PowerPoint.SlideShowWindow.View.md)property of the  **SlideShowWindow** object to return the **SlideShowView** object. The following example sets slide show window one to display the first slide in the presentation.


```vb
SlideShowWindows(1).View.First
```

Use the [Run](PowerPoint.SlideShowSettings.Run.md)method of the  **SlideShowSettings** object to create a **SlideShowWindow** object, and then use the **View** property to return the **SlideShowView** object the window contains. The following example runs a slide show of the active presentation, changes the pointer to a pen, and sets the pen color for the slide show to red.




```vb
With ActivePresentation.SlideShowSettings.Run.View

    .PointerColor.RGB = RGB(255, 0, 0)

    .PointerType = ppSlideShowPointerPen

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)
[SlideShowView Object Members](overview/PowerPoint.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]