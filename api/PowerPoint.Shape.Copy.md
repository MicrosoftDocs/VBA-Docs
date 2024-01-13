---
title: Shape.Copy method (PowerPoint)
keywords: vbapp10.chm547051
f1_keywords:
- vbapp10.chm547051
api_name:
- PowerPoint.Shape.Copy
ms.assetid: 41c82fd1-9ee7-c937-0a75-77b84c33c972
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Shape.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Remarks

Use the **Paste** method to paste the contents of the Clipboard.

If the shape is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example copies shapes one and two on slide one in the active presentation to the Clipboard and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Copy

    .Slides(2).Shapes.Paste

End With
```


## See also


[Shape Object](PowerPoint.Shape.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]