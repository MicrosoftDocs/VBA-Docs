---
title: PageSetup.PageWidth property (Publisher)
keywords: vbapb10.chm6946822
f1_keywords:
- vbapb10.chm6946822
ms.prod: publisher
api_name:
- Publisher.PageSetup.PageWidth
ms.assetid: 547f2881-d9fa-fa5f-1643-ab08084fb423
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.PageWidth property (Publisher)

Returns or sets a **Variant** that represents the width of the pages in a publication. Read/write.


## Syntax

_expression_.**PageWidth**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated as [points](../language/glossary/vbe-glossary.md#point). String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet width and the page width.


## Example

The following example sets a width of eight inches for the pages in the active publication.

```vb
Public Sub PageWidth_Example() 
 ActiveDocument.PageSetup.PageWidth = InchesToPoints(8) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]