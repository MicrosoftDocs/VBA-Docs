---
title: View.PrintOptions property (PowerPoint)
keywords: vbapp10.chm512011
f1_keywords:
- vbapp10.chm512011
ms.prod: powerpoint
api_name:
- PowerPoint.View.PrintOptions
ms.assetid: ee0aeece-e1f9-36ce-1d5d-cec9e0e4046b
ms.date: 06/08/2017
localization_priority: Normal
---


# View.PrintOptions property (PowerPoint)

Returns a **[PrintOptions](PowerPoint.PrintOptions.md)** object that represents print options that are saved with the specified presentation. Read-only.


## Syntax

_expression_. `PrintOptions`

_expression_ A variable that represents a [View](PowerPoint.View.md) object.


## Return value

PrintOptions


## Example

This example causes hidden slides in the active presentation to be printed, and it scales the printed slides to fit the paper size.


```vb
With Application.ActivePresentation

    With .PrintOptions

        .PrintHiddenSlides = True

        .FitToPage = True

    End With

    .PrintOut

End With
```


## See also


[View Object](PowerPoint.View.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]