---
title: SlideRange.Cut method (PowerPoint)
keywords: vbapp10.chm532012
f1_keywords:
- vbapp10.chm532012
api_name:
- PowerPoint.SlideRange.Cut
ms.assetid: 91d80a2b-e67a-290b-cb41-6bbeeb467d1b
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# SlideRange.Cut method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Remarks

If any slide in the range is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example deletes shapes one and two from slide one in the active presentation, places copies of them on the Clipboard, and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Cut

    .Slides(2).Shapes.Paste

End With
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]