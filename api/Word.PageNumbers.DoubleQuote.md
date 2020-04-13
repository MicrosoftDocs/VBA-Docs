---
title: PageNumbers.DoubleQuote property (Word)
keywords: vbawd10.chm159776778
f1_keywords:
- vbawd10.chm159776778
ms.prod: word
api_name:
- Word.PageNumbers.DoubleQuote
ms.assetid: 38a63f94-2a47-cea5-69a8-16616458fb9a
ms.date: 06/08/2017
localization_priority: Normal
---


# PageNumbers.DoubleQuote property (Word)

 **True** if Microsoft Word encloses the specified **PageNumbers** object in double quotation marks ("). Read/write **Boolean**.


## Syntax

_expression_. `DoubleQuote`

 _expression_ An expression that returns a '[PageNumbers](Word.pagenumbers.md)' object.


## Remarks

To set Word to enclose page numbers in double quotation marks by default, use the **[AddHebDoubleQuote](Word.Options.AddHebDoubleQuote.md)** property.


## Example

This example encloses the page numbers in the first footer of the active document in double quotation marks (").


```vb
ActiveDocument.Sections(1).Footers(1) _ 
 .PageNumbers.DoubleQuote = True
```


## See also


[PageNumbers Collection Object](Word.pagenumbers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]