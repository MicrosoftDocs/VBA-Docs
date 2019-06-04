---
title: Application.AfterPresentationOpen event (PowerPoint)
keywords: vbapp10.chm621021
f1_keywords:
- vbapp10.chm621021
ms.prod: powerpoint
api_name:
- PowerPoint.Application.AfterPresentationOpen
ms.assetid: 3f783486-0ceb-166d-017b-0a41bd15cfa6
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AfterPresentationOpen event (PowerPoint)

Occurs after an existing presentation is opened.


## Syntax

_expression_. `AfterPresentationOpen`( `_Pres_` )

 _expression_ An expression that returns an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Pres_|Required|**Presentation**|The presentation that is opened.|

## Example

This example modifies the background color for color scheme three, applies the modified color scheme to the presentation that was opened, and displays the presentation in Slide view.


```vb
Private Sub App_AfterPresentationOpen(ByVal Pres As Presentation)

    With Pres

        Set CS3 = .ColorSchemes(3)

        CS3.Colors(ppBackground).RGB = RGB(240, 115, 100)

        With Windows(1)

            .Selection.SlideRange.ColorScheme = CS3

            .ViewType = ppViewSlide

        End With

    End With

End Sub
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]