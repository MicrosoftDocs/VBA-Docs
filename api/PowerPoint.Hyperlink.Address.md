---
title: Hyperlink.Address property (PowerPoint)
keywords: vbapp10.chm526004
f1_keywords:
- vbapp10.chm526004
ms.prod: powerpoint
api_name:
- PowerPoint.Hyperlink.Address
ms.assetid: d3d2174a-fbb2-432d-bc42-6623c91e9843
ms.date: 06/08/2017
localization_priority: Normal
---


# Hyperlink.Address property (PowerPoint)

Returns or sets the Internet address (URL) to the target document. Read/write.


## Syntax

_expression_.**Address**

_expression_ A variable that represents an [Hyperlink](PowerPoint.Hyperlink.md) object.


## Return value

String


## Example

This example scans all shapes on the first slide for the URL to the Microsoft Web site.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Hyperlinks

    If s.Address = "https://www.microsoft.com/" Then

        MsgBox "You have a link to the Microsoft Home Page"

    End If

Next
```


## See also


[Hyperlink Object](PowerPoint.Hyperlink.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]