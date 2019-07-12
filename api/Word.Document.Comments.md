---
title: Document.Comments property (Word)
keywords: vbawd10.chm158007305
f1_keywords:
- vbawd10.chm158007305
ms.prod: word
api_name:
- Word.Document.Comments
ms.assetid: 1597a002-afa4-743d-60a6-ffd398f2b599
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Comments property (Word)

Returns a **[Comments](Word.comments.md)** collection that represents all the comments in the specified document. Read-only.


## Syntax

_expression_.**Comments**

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example compares the author name of each comment in the active document with the user name on the **User Information** tab in the **Options** dialog box (**Tools** menu). If the names aren't the same, the comment reference mark is formatted to appear in red.

```vb
For Each comm In ActiveDocument.Comments 
 If comm.Author <> Application.UserName Then _ 
 comm.Reference.Font.ColorIndex = wdRed 
Next comm
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]