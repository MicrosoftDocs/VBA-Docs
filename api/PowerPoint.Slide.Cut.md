---
title: Slide.Cut method (PowerPoint)
keywords: vbapp10.chm531012
f1_keywords:
- vbapp10.chm531012
api_name:
- PowerPoint.Slide.Cut
ms.assetid: 03029017-52c8-5176-a218-8b5ff8edec10
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Slide.Cut method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a [Slide](PowerPoint.Slide.md) object.


## Remarks

If the slide is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example deletes slide one from the active presentation and places a copy of it on the Clipboard.


```vb
ActivePresentation.Slides(1).Cut
```


## See also


[Slide Object](PowerPoint.Slide.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]