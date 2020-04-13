---
title: Presentation.Fonts property (PowerPoint)
keywords: vbapp10.chm583016
f1_keywords:
- vbapp10.chm583016
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Fonts
ms.assetid: 3caece78-6ca9-bca8-5683-4722e1f563cf
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.Fonts property (PowerPoint)

Returns a **[Fonts](PowerPoint.Fonts.md)** collection that represents all fonts used in the specified presentation. Read-only.


## Syntax

_expression_.**Fonts**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Example

This example replaces the Times New Roman font with the Courier font in the active presentation.


```vb
Application.ActivePresentation.Fonts _
    .Replace "Times New Roman", "Courier"
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]