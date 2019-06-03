---
title: Range.Frames property (Word)
keywords: vbawd10.chm157155394
f1_keywords:
- vbawd10.chm157155394
ms.prod: word
api_name:
- Word.Range.Frames
ms.assetid: c30bb71d-3998-42fe-2850-a76c3975418b
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Frames property (Word)

Returns a  **[Frames](Word.Frames.md)** collection that represents all the frames in a range. Read-only.


## Syntax

_expression_. `Frames`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example causes text to wrap around frames in the first section in the active document.


```vb
For Each aFrame In ActiveDocument.Sections(1).Range.Frames 
 aFrame.TextWrap = True 
Next aFrame
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]