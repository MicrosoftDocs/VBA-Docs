---
title: SlideRange.Copy method (PowerPoint)
keywords: vbapp10.chm532013
f1_keywords:
- vbapp10.chm532013
api_name:
- PowerPoint.SlideRange.Copy
ms.assetid: d781370d-8107-efaa-77ea-a7f1aa58737b
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# SlideRange.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [SlideRange](PowerPoint.SlideRange.md) object.


## Remarks

Use the **Paste** method to paste the contents of the Clipboard.

If any slide in the range is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example copies slide one in the active presentation to the Clipboard.


```vb
ActivePresentation.Slides(1).Copy
```


## See also


[SlideRange Object](PowerPoint.SlideRange.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]