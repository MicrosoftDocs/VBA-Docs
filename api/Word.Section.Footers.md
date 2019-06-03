---
title: Section.Footers property (Word)
keywords: vbawd10.chm156827770
f1_keywords:
- vbawd10.chm156827770
ms.prod: word
api_name:
- Word.Section.Footers
ms.assetid: 2aa522ae-fc34-eb75-790f-85a8182f76c2
ms.date: 06/08/2017
localization_priority: Normal
---


# Section.Footers property (Word)

Returns a  **[HeadersFooters](Word.headersfooters.md)** collection that represents the footers in the specified section. Read-only.


## Syntax

_expression_. `Footers`

_expression_ A variable that represents a '[Section](Word.Section.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md). To return a  **HeadersFooters** collection that represents the headers for the specified section, use the **[Headers](Word.Section.Headers.md)** property.


## Example

This example adds a right-aligned page number to the primary footer in the first section in the active document.


```vb
With ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) 
 .PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight 
End With
```


## See also


[Section Object](Word.Section.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]