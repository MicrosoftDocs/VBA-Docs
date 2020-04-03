---
title: ParagraphFormat.TabStops property (Word)
keywords: vbawd10.chm156435535
f1_keywords:
- vbawd10.chm156435535
ms.prod: word
api_name:
- Word.ParagraphFormat.TabStops
ms.assetid: 9eed85b9-aee6-04af-c5ce-f6ba47176d35
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.TabStops property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.


## Syntax

_expression_. `TabStops`

_expression_ A variable that represents a '[ParagraphFormat](Word.ParagraphFormat.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example adds a centered tab stop at 2 inches to all the paragraphs in the active document. The  **InchesToPoints** method is used to convert inches to points.


```vb
With ActiveDocument.Paragraphs.TabStops 
 .Add Position:= InchesToPoints(2), Alignment:= wdAlignTabCenter 
End With
```

This example sets the tab stops for every paragraph in the document to match the tab stops in the first paragraph.




```vb
Set para1Tabs = ActiveDocument.Paragraphs(1).TabStops 
ActiveDocument.Paragraphs.TabStops = para1Tabs
```


## See also


[ParagraphFormat Object](Word.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]