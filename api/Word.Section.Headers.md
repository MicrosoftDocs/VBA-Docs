---
title: Section.Headers property (Word)
keywords: vbawd10.chm156827769
f1_keywords:
- vbawd10.chm156827769
ms.prod: word
api_name:
- Word.Section.Headers
ms.assetid: 72b61449-2f93-a67a-2757-3c0441961307
ms.date: 06/08/2017
localization_priority: Normal
---


# Section.Headers property (Word)

Returns a  **[HeadersFooters](Word.headersfooters.md)** collection that represents the headers for the specified section. Read-only.


## Syntax

_expression_. `Headers`

_expression_ A variable that represents a '[Section](Word.Section.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md). To return a  **HeadersFooters** collection that represents the footers for the specified section, use the **[Footers](Word.Section.Footers.md)** property.


## Example

This example adds centered page numbers to every page in the active document except the first. (A separate header is created for the first page.)


```vb
With ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) 
 .PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberCenter, _ 
 FirstPage:=False 
End With
```

This example adds text to the first-page header in the active document.




```vb
ActiveDocument.PageSetup.DifferentFirstPageHeaderFooter = True 
With ActiveDocument.Sections(1).Headers(wdHeaderFooterFirstPage) 
 .Range.InsertAfter("First Page Text") 
 .Range.Paragraphs.Alignment = wdAlignParagraphRight 
End With
```


## See also


[Section Object](Word.Section.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
