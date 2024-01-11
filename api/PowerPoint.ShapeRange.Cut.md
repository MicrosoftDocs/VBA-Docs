---
title: ShapeRange.Cut method (PowerPoint)
keywords: vbapp10.chm548050
f1_keywords:
- vbapp10.chm548050
api_name:
- PowerPoint.ShapeRange.Cut
ms.assetid: 0e86d67c-7d52-4f3a-4cdd-6363667600a1
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# ShapeRange.Cut method (PowerPoint)

Deletes the specified object and places it on the Clipboard.


## Syntax

_expression_.**Cut**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Remarks

If any shape in the range is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example deletes shapes one and two from slide one in the active presentation, places copies of them on the Clipboard, and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Cut

    .Slides(2).Shapes.Paste

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]