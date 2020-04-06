---
title: Paragraphs.TabStops property (Word)
keywords: vbawd10.chm156763215
f1_keywords:
- vbawd10.chm156763215
ms.prod: word
api_name:
- Word.Paragraphs.TabStops
ms.assetid: cf369030-7569-699f-d8be-7a24b63e22eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.TabStops property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraphs. Read/write.


## Syntax

_expression_. `TabStops`

_expression_ A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


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


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]