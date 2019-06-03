---
title: Range.Words property (Word)
keywords: vbawd10.chm157155379
f1_keywords:
- vbawd10.chm157155379
ms.prod: word
api_name:
- Word.Range.Words
ms.assetid: ada98916-b87c-7592-ee2d-561ed7067f39
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Words property (Word)

Returns a  **Words** collection that represents all the words in a range. Read-only.


## Syntax

_expression_. `Words`

_expression_ A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

Punctuation and paragraph marks in a document are included in the  **Words** collection.

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example displays the number of words in the selection. Paragraphs marks, partial words, and punctuation are included in the count.


```vb
MsgBox "There are " & Selection.Words.Count & " words."
```

This example steps through the words in  _myRange_ (which spans from the beginning of the active document to the end of the selection) and deletes the word "Franklin" (including the trailing space) wherever it occurs in the range.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=Selection.End) 
For Each aWord In myRange.Words 
 If aWord.Text = "Franklin " Then aWord.Delete 
Next aWord
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]