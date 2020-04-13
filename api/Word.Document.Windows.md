---
title: Document.Windows property (Word)
keywords: vbawd10.chm158007330
f1_keywords:
- vbawd10.chm158007330
ms.prod: word
api_name:
- Word.Document.Windows
ms.assetid: bb075fd7-2dae-18c9-f49a-0c478d840b76
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Windows property (Word)

Returns a  **[Windows](Word.windows.md)** collection that represents all windows for the specified document. Read-only.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of windows for the active document, both before and after the **NewWindow** method is run.


```vb
MsgBox Prompt:= ActiveDocument.Windows.Count & " window(s)", _ 
 Title:= ActiveDocument.Name 
ActiveDocument.ActiveWindow.NewWindow 
MsgBox Prompt:= ActiveDocument.Windows.Count & " windows", _ 
 Title:= ActiveDocument.Name
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]