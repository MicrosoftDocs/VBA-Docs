---
title: HeaderFooter.Exists property (Word)
keywords: vbawd10.chm159711236
f1_keywords:
- vbawd10.chm159711236
ms.prod: word
api_name:
- Word.HeaderFooter.Exists
ms.assetid: 84ce3ac9-a4be-f99a-eb4b-1a145373659f
ms.date: 06/08/2017
localization_priority: Normal
---


# HeaderFooter.Exists property (Word)

 **True** if the specified **HeaderFooter** object exists. Read/write **Boolean**.


## Syntax

_expression_. `Exists`

_expression_ A variable that represents a '[HeaderFooter](Word.HeaderFooter.md)' object.


## Remarks

The primary header and footer exist in all new documents by default. Use this method to determine whether a first-page or odd-page header or footer exists. You can also use the  **[DifferentFirstPageHeaderFooter](Word.PageSetup.DifferentFirstPageHeaderFooter.md)** or **[OddAndEvenPagesHeaderFooter](Word.PageSetup.OddAndEvenPagesHeaderFooter.md)** property to return or set the number of headers and footers in the specified document or section.


## Example

If a first-page header exists in section one, this example sets the text for the header.


```vb
Dim secTemp As Section 
 
Set secTemp = ActiveDocument.Sections(1) 
If secTemp.Headers(wdHeaderFooterFirstPage).Exists = True Then 
 secTemp.Headers(wdHeaderFooterFirstPage).Range.Text = _ 
 "First Page" 
End If
```


## See also


[HeaderFooter Object](Word.HeaderFooter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]