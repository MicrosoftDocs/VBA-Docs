---
title: Paragraph.TabStops property (Word)
keywords: vbawd10.chm156697679
f1_keywords:
- vbawd10.chm156697679
ms.prod: word
api_name:
- Word.Paragraph.TabStops
ms.assetid: e1739724-c236-e934-4e10-512d19cb8989
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph.TabStops property (Word)

Returns or sets a  **TabStops** collection that represents all the custom tab stops for the specified paragraph. Read/write.


## Syntax

_expression_. `TabStops`

_expression_ A variable that represents a '[Paragraph](Word.Paragraph.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example sets the tab stops for every paragraph in the document to match the tab stops in the first paragraph.


```vb
Set para1Tabs = ActiveDocument.Paragraphs(1).TabStops 
ActiveDocument.Paragraphs.TabStops = para1Tabs
```


## See also


[Paragraph Object](Word.Paragraph.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]