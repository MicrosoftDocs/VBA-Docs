---
title: Borders.ApplyPageBordersToAllSections method (Word)
keywords: vbawd10.chm154929104
f1_keywords:
- vbawd10.chm154929104
ms.prod: word
api_name:
- Word.Borders.ApplyPageBordersToAllSections
ms.assetid: 03905cb9-89f6-9bfa-65a1-dd23880e9c23
ms.date: 06/08/2017
localization_priority: Normal
---


# Borders.ApplyPageBordersToAllSections method (Word)

Applies the specified page-border formatting to all sections in a document.


## Syntax

 _expression_. `ApplyPageBordersToAllSections`

 _expression_ Required. A variable that represents a '[Borders](Word.borders.md)' collection.


## Example

This example adds a single-line page border to all sections in the active document.


```vb
Dim borderLoop As Border 
 
With ActiveDocument.Sections(1) 
 For Each borderLoop In .Borders 
 With borderLoop 
 .LineStyle = wdLineStyleSingle 
 .LineWidth = wdLineWidth050pt 
 End With 
 Next borderLoop 
 .Borders.ApplyPageBordersToAllSections 
End With
```


## See also


[Borders Collection Object](Word.borders.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]