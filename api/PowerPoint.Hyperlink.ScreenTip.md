---
title: Hyperlink.ScreenTip property (PowerPoint)
keywords: vbapp10.chm526008
f1_keywords:
- vbapp10.chm526008
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.ScreenTip
ms.assetid: 96ff1076-7563-8250-ea75-cee46094824e
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.ScreenTip property (PowerPoint)

Returns or sets the ScreenTip text of a hyperlink. Read/write.


## Syntax

_expression_.**ScreenTip**

_expression_ A variable that represents a [Hyperlink](PowerPoint.Hyperlink.md) object.


## Return value

String


## Remarks

ScreenTip text appears, for example, when you save a presentation to HTML, view it in a web browser, and rest the mouse pointer over a hyperlink. Some browsers may not support ScreenTips.


## Example

This example sets the ScreenTip text for the first hyperlink.


```vb
ActivePresentation.Slides(1).Hyperlinks(1) _
    .ScreenTip = "Go to the Microsoft home page"
```


## See also


[Hyperlink Object](PowerPoint.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]