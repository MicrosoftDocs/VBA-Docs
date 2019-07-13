---
title: SlideShowView.SlideShowName property (PowerPoint)
keywords: vbapp10.chm513014
f1_keywords:
- vbapp10.chm513014
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.SlideShowName
ms.assetid: 63efa2d8-7321-dc72-3c25-ab5ab4ba5c0a
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.SlideShowName property (PowerPoint)

Returns the name of the custom slide show that's currently running in the specified slide show view. Read-only.


## Syntax

_expression_. `SlideShowName`

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Return value

String


## Example

If the slide show running in slide show window one is a custom slide show, this example displays its name.


```vb
With SlideShowWindows(1).View
    If .IsNamedShow Then
        MsgBox "Now showing in slide show window 1: " _
            & .SlideShowName
    End If
End With
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]