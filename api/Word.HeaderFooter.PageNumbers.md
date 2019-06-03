---
title: HeaderFooter.PageNumbers property (Word)
keywords: vbawd10.chm159711237
f1_keywords:
- vbawd10.chm159711237
ms.prod: word
api_name:
- Word.HeaderFooter.PageNumbers
ms.assetid: 2e36c668-f696-e09e-dd04-ae77e7524232
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter.PageNumbers property (Word)

Returns a  **[PageNumbers](Word.pagenumbers.md)** collection that represents all the page number fields included in the specified header or footer.


## Syntax

_expression_. `PageNumbers`

 _expression_ An expression that returns a '[HeaderFooter](Word.HeaderFooter.md)' object.


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../word/Concepts/Miscellaneous/returning-an-object-from-a-collection-word.md).


## Example

This example creates a new document and adds page numbers to the footer.


```vb
Set myDoc = Documents.Add 
With myDoc.Sections(1).Footers(wdHeaderFooterPrimary) 
 .PageNumbers.Add PageNumberAlignment := wdAlignPageNumberCenter 
End With
```


## See also


[HeaderFooter Object](Word.HeaderFooter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]