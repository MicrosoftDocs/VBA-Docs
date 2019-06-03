---
title: Selection.Frames property (Word)
keywords: vbawd10.chm158662722
f1_keywords:
- vbawd10.chm158662722
ms.prod: word
api_name:
- Word.Selection.Frames
ms.assetid: cc589559-858a-2ebb-00dd-64f97966859f
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Frames property (Word)

Returns a  **[Frames](Word.Frames.md)** collection that represents all the frames in a selection. Read-only.


## Syntax

_expression_. `Frames`

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example causes text to wrap around frames in the first section in the active document.


```vb
For Each aFrame In ActiveDocument.Sections(1).Range.Frames 
 aFrame.TextWrap = True 
Next aFrame
```

This example adds a frame around the selection and returns a frame object to the myFrame variable.




```vb
Set myFrame = ActiveDocument.Frames.Add(Range:=Selection.Range)
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]