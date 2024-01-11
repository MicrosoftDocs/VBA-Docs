---
title: Selection.Copy method (PowerPoint)
keywords: vbapp10.chm508004
f1_keywords:
- vbapp10.chm508004
api_name:
- PowerPoint.Selection.Copy
ms.assetid: 954106da-a2a9-0c55-114a-5a79f578e0c4
ms.date: 08/02/2022
ms.localizationpriority: medium
---


# Selection.Copy method (PowerPoint)

Copies the specified object to the Clipboard.


## Syntax

_expression_.**Copy**

_expression_ A variable that represents a [Selection](PowerPoint.Selection.md) object.


## Remarks

Use the **Paste** method to paste the contents of the Clipboard.

If the selected content is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).


## Example

This example copies the selection in window one to the Clipboard and then pastes it into the view in window two. If the Clipboard contents cannot be pasted into the view in window two — for example, if you try to paste a shape into slide sorter view — this example fails.


```vb
Windows(1).Selection.Copy

Windows(2).View.Paste
```


## See also


[Selection Object](PowerPoint.Selection.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]