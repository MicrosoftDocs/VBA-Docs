---
title: TextRange.InsertSlideNumber method (PowerPoint)
keywords: vbapp10.chm569021
f1_keywords:
- vbapp10.chm569021
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.InsertSlideNumber
ms.assetid: 07489db8-9db1-9721-845a-7895ad316aca
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.InsertSlideNumber method (PowerPoint)

Inserts the slide number of the current slide into the specified text range. Returns a  **TextRange** object that represents the slide number.


## Syntax

_expression_. `InsertSlideNumber`

_expression_ A variable that represents an [TextRange](PowerPoint.TextRange.md) object.


## Return value

TextRange


## Remarks

The inserted slide number is automatically updated when the slide number of the current slide changes.


## Example

This example inserts the slide number of the current slide after the first sentence of the first paragraph in shape two on slide one in the active presentation.


```vb
Set sh = Application.ActivePresentation.Slides(1).Shapes(2)

Set sentOne = sh.TextFrame.TextRange.Paragraphs(1).Sentences(1)

sentOne.InsertAfter.InsertSlideNumber
```


## See also


[TextRange Object](PowerPoint.TextRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]