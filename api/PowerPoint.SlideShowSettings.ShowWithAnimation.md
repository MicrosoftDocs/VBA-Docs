---
title: SlideShowSettings.ShowWithAnimation property (PowerPoint)
keywords: vbapp10.chm514012
f1_keywords:
- vbapp10.chm514012
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.ShowWithAnimation
ms.assetid: 9255fc7b-50fa-c65e-5ef4-3c214dede4a4
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowSettings.ShowWithAnimation property (PowerPoint)

Determines whether the specified slide show displays shapes with assigned animation settings. Read/write.


## Syntax

_expression_. `ShowWithAnimation`

_expression_ A variable that represents a [SlideShowSettings](PowerPoint.SlideShowSettings.md) object.


## Return value

MsoTriState


## Remarks

The value of the  **ShowWithAnimation** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified slide show displays shapes without assigned animation settings.|
|**msoTrue**| The specified slide show displays shapes with assigned animation settings.|

## Example

This example runs a slide show of the active presentation with animation and narration turned off.


```vb
With ActivePresentation.SlideShowSettings

    .ShowWithAnimation = msoFalse

    .ShowWithNarration = msoFalse

    .Run

End With
```


## See also


[SlideShowSettings Object](PowerPoint.SlideShowSettings.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]