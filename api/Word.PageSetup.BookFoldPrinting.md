---
title: PageSetup.BookFoldPrinting property (Word)
keywords: vbawd10.chm158401735
f1_keywords:
- vbawd10.chm158401735
ms.prod: word
api_name:
- Word.PageSetup.BookFoldPrinting
ms.assetid: 05bed9bc-5a93-9978-6faf-3fbc6d7239a3
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.BookFoldPrinting property (Word)

**True** for Microsoft Word to print a document in a series of booklets so that the printed pages can be folded and read as a book. Read/write **Boolean**.


## Syntax

_expression_.**BookFoldPrinting**

_expression_ An expression that returns a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example turns the active document into a booklet that prints in four-page increments.

```vb
Sub Booklet() 
 With PageSetup 
 .BookFoldPrinting = True 
 .BookFoldPrintingSheets = 4 
 End With 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]