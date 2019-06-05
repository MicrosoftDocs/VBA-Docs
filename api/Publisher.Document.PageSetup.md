---
title: Document.PageSetup property (Publisher)
keywords: vbapb10.chm196632
f1_keywords:
- vbapb10.chm196632
ms.prod: publisher
api_name:
- Publisher.Document.PageSetup
ms.assetid: 1dac39f0-2507-a85b-8c71-cd1980022fb3
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.PageSetup property (Publisher)

Returns a **[PageSetup](Publisher.PageSetup.md)** object representing a publication's page size, page layout, and paper settings. Read-only.


## Syntax

_expression_.**PageSetup**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

PageSetup


## Remarks

You can only use the **PageSetup** property when printing multiple pages on a single sheet of printer paper. If the page size is greater than half the paper size, Microsoft Publisher displays an error.


## Example

This example specifies page setup options for a publication with multiple publication pages printed on each sheet of printer paper.

```vb
Sub SetTopMargin() 
 With ActiveDocument.PageSetup 
 .PageHeight = InchesToPoints(5) 
 .PageWidth = InchesToPoints(8) 
 .MultiplePagesPerSheet = True 
 .TopMargin = InchesToPoints(0.25) 
 .LeftMargin = InchesToPoints(0.25) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]