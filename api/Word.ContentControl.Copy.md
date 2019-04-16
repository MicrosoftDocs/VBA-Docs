---
title: ContentControl.Copy method (Word)
keywords: vbawd10.chm266534918
f1_keywords:
- vbawd10.chm266534918
ms.prod: word
api_name:
- Word.ContentControl.Copy
ms.assetid: ce3ba4ce-aef7-cb7d-ec7b-a160155a939d
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControl.Copy method (Word)

Copies the content control from the active document to the Clipboard.


## Syntax

_expression_.**Copy**

 _expression_ An expression that returns a [ContentControl](./Word.ContentControl.md) object.


## Remarks

When you use the  **Copy** method, the original content control remains in the active document, but a copy of the control, including all text and property settings, is moved to the Clipboard. You can then paste the content control into other sections of the active document. Use the **Paste** method of the **[Selection](Word.Selection.md)** object or the **Paste** method of the **[Range](Word.Range.md)** object to insert the copied content control, or use the **Paste** function from within Microsoft Word.


## See also


[ContentControl Object](Word.ContentControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]