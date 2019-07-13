---
title: View.GotoSlide method (PowerPoint)
keywords: vbapp10.chm512007
f1_keywords:
- vbapp10.chm512007
ms.prod: powerpoint
api_name:
- PowerPoint.View.GotoSlide
ms.assetid: bb898aa7-d2b5-0728-90dd-2f4ce399bb21
ms.date: 06/08/2017
localization_priority: Normal
---


# View.GotoSlide method (PowerPoint)

Switches to the specified slide.


## Syntax

_expression_. `GotoSlide` (_Index_)

_expression_ A variable that represents a [View](PowerPoint.View.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The number of the slide to switch to.|

## Example

This example switches from the current slide to the slide three in slide show window one. If you switch back to the current slide during the slide show, its entire animation will start over.


```vb
With SlideShowWindows(1).View

    .GotoSlide 3

End With
```

This example switches from the current slide to the slide three in slide show window one. If you switch back to the current slide during the slide show, its animation will pick up where it left off.




```vb
With SlideShowWindows(1).View

    .GotoSlide 3, msoFalse

End With
```


## See also


[View Object](PowerPoint.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
