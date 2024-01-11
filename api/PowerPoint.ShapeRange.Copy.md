---
title: ShapeRange.Copy method (PowerPoint)
keywords: vbapp10.chm548051
f1_keywords:
- vbapp10.chm548051
api_name:
- PowerPoint.ShapeRange.Copy
ms.assetid: ddc0dad9-6647-e2f4-393a-347c273656dd
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# ShapeRange.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Remarks

Use the **Paste** method to paste the contents of the Clipboard.

If any shape in the range is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example copies shapes one and two on slide one in the active presentation to the Clipboard and then pastes the copies onto slide two.


```vb
With ActivePresentation

    .Slides(1).Shapes.Range(Array(1, 2)).Copy

    .Slides(2).Shapes.Paste

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]