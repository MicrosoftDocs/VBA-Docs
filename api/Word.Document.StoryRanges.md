---
title: Document.StoryRanges property (Word)
keywords: vbawd10.chm158007352
f1_keywords:
- vbawd10.chm158007352
ms.prod: word
api_name:
- Word.Document.StoryRanges
ms.assetid: 6afc9e1a-950c-e1b0-15d5-73afeb72fc59
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.StoryRanges property (Word)

Returns a  **[StoryRanges](Word.storyranges.md)** collection that represents all the stories in the specified document. Read-only.


## Syntax

_expression_. `StoryRanges`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example steps through the **StoryRanges** collection to determine whether **wdPrimaryFooterStory** is part of the **StoryRanges** collection.


```vb
For Each aStory In ActiveDocument.StoryRanges 
 If aStory.StoryType = wdEvenPagesFooterStory Then 
 MsgBox "Document includes an even page footer" 
 End If 
Next aStory
```

This example adds text to the primary header story and then displays the text.




```vb
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary).Range _ 
 .Text = "Header text" 
MsgBox ActiveDocument.StoryRanges(wdPrimaryHeaderStory).Text
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]