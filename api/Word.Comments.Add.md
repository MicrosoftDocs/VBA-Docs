---
title: Comments.Add method (Word)
keywords: vbawd10.chm155189252
f1_keywords:
- vbawd10.chm155189252
ms.prod: word
api_name:
- Word.Comments.Add
ms.assetid: bf3e2f9b-b7d6-f669-c82a-70ff58aaedfe
ms.date: 06/08/2017
localization_priority: Normal
---


# Comments.Add method (Word)

Returns a  **Comment** object that represents a comment added to a range.


## Syntax

_expression_.**Add** (_Range_, _Text_)

_expression_ Required. A variable that represents a '[Comments](Word.comments.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range to have a comment added to it.|
| _Text_|Optional| **Variant**|The text of the comment.|

## Return value

Comment


## Example

This example adds a comment at the insertion point.


```vb
Sub AddComment() 
 Selection.Collapse Direction:=wdCollapseEnd 
 ActiveDocument.Comments.Add _ 
 Range:=Selection.Range, Text:="review this" 
End Sub
```

This example adds a comment to the third paragraph in the active document.




```vb
Sub Comment3rd() 
 Dim myRange As Range 
 
 Set myRange = ActiveDocument.Paragraphs(3).Range 
 ActiveDocument.Comments.Add Range:=myRange, _ 
 Text:="original third paragraph" 
End Sub
```


## See also


[Comments Collection Object](Word.comments.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]