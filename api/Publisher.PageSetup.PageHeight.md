---
title: PageSetup.PageHeight property (Publisher)
keywords: vbapb10.chm6946821
f1_keywords:
- vbapb10.chm6946821
ms.prod: publisher
api_name:
- Publisher.PageSetup.PageHeight
ms.assetid: 1ef153e2-5d13-d896-cd69-2066efa2f8ef
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.PageHeight property (Publisher)

Returns or sets a **Variant** that represents the height of the pages in a publication. Read/write.


## Syntax

_expression_.**PageHeight**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated as [points](../language/glossary/vbe-glossary.md#point). String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet height and the page height.


## Example

This example specifies a height of five inches for the pages in the active publication.

```vb
Public Sub PageHeight_Example() 
 ActiveDocument.PageSetup.PageHeight = InchesToPoints(5) 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]