---
title: ContentControl.Range property (Word)
keywords: vbawd10.chm266534913
f1_keywords:
- vbawd10.chm266534913
ms.prod: word
api_name:
- Word.ContentControl.Range
ms.assetid: e83efa5d-edd7-2cdc-ee6f-925db82e3d40
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControl.Range property (Word)

Returns a  **[Range](Word.Range.md)** that represents the contents of the content control in the active document. Read-only.


## Syntax

_expression_.**Range**

 _expression_ An expression that returns a [ContentControl](./Word.ContentControl.md) object.


## Remarks

Use the  **Range** property to access the text, the formatting for the text, and other text properties. For more information, see [Working with Range Objects](../word/Concepts/Working-with-Word/working-with-range-objects.md).


## Example

The following example inserts a date content control into the active document, and then sets the contents of the content control and specifies that the user cannot edit the contents or delete the control from the document.


```vb
Dim objCC As ContentControl 
 
Set objCC = ActiveDocument.ContentControls _ 
 .Add(wdContentControlDate) 
 
objCC.Range.Text = "January 1, 2007" 
objCC.LockContents = True 
objCC.LockContentControl = True
```


## See also


[ContentControl Object](Word.ContentControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]