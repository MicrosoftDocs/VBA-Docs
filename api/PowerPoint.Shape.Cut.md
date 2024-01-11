---
title: Shape.Cut method (PowerPoint)
keywords: vbapp10.chm547050
f1_keywords:
- vbapp10.chm547050
api_name:
- PowerPoint.Shape.Cut
ms.assetid: 908c998d-a15f-5075-33e0-de6c124a0ec7
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Shape.Cut method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Remarks

If the shape is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example deletes shapes one and two from slide one in the active presentation, places copies of them on the Clipboard, and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Cut

    .Slides(2).Shapes.Paste

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]