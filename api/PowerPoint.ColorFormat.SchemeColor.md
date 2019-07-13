---
title: ColorFormat.SchemeColor property (PowerPoint)
keywords: vbapp10.chm506004
f1_keywords:
- vbapp10.chm506004
ms.prod: powerpoint
api_name:
- PowerPoint.ColorFormat.SchemeColor
ms.assetid: 4c62e7a7-ce53-c93e-9ec5-b299a18f5bf7
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorFormat.SchemeColor property (PowerPoint)

Returns or sets the color in the applied color scheme that's associated with the specified object. Read/write.


## Syntax

_expression_. `SchemeColor`

_expression_ A variable that represents a [ColorFormat](PowerPoint.ColorFormat.md) object.


## Return value

PpColorSchemeIndex


## Remarks

The value of the  **SchemeColor** property can be one of these **PpColorSchemeIndex** constants.


||
|:-----|
|**ppAccent1**|
|**ppAccent2**|
|**ppAccent3**|
|**ppBackground**|
|**ppFill**|
|**ppForeground**|
|**ppNotSchemeColor**|
|**ppSchemeColorMixed**|
|**ppShadow**|
|**ppTitle**|

## Example

This example switches the background color on slide one in the active presentation between an explicit red-green-blue value and the color-scheme background color.


```vb
With ActivePresentation.Slides(1)

    .FollowMasterBackground = False

    With .Background.Fill.ForeColor

        If .Type = msoColorTypeScheme Then

            .RGB = RGB(0, 128, 128)

        Else

            .SchemeColor = ppBackground

        End If

    End With

End With
```


## See also


[ColorFormat Object](PowerPoint.ColorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]