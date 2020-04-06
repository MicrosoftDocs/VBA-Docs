---
title: View.Slide property (PowerPoint)
keywords: vbapp10.chm512006
f1_keywords:
- vbapp10.chm512006
ms.prod: powerpoint
api_name:
- PowerPoint.View.Slide
ms.assetid: 18a2f9e0-ae3d-b662-90d4-a0c0de18d073
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Slide property (PowerPoint)

Returns or sets a **[Slide](PowerPoint.Slide.md)** object that represents the slide that's currently displayed in the specified document window view. Read/write.


## Syntax

_expression_. `Slide`

_expression_ A variable that represents a [View](PowerPoint.View.md) object.


## Remarks

If the currently displayed slide is from an embedded presentation, you can use the  **[Parent](PowerPoint.Slide.Parent.md)** property of the **Slide** object returned by the **Slide** property to return the embedded presentation that contains the slide. (The **[Presentation](PowerPoint.SlideShowWindow.Presentation.md)** property of the **SlideShowWindow** object or **DocumentWindow** object returns the presentation from which the window was created, not the embedded presentation.)


## Example

This example places on the Clipboard a copy of the slide that's currently displayed in slide show window one.


```vb
SlideShowWindows(1).View.Slide.Copy
```

This example displays the name of the presentation currently running in slide show window one.




```vb
MsgBox SlideShowWindows(1).View.Slide.Parent.Name
```


## See also


[View Object](PowerPoint.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]