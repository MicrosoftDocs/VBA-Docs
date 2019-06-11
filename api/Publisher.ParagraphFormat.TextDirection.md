---
title: ParagraphFormat.TextDirection property (Publisher)
keywords: vbapb10.chm5439507
f1_keywords:
- vbapb10.chm5439507
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.TextDirection
ms.assetid: b96c634d-0e7e-dba8-2bf4-e5baf3afa3d1
ms.date: 06/12/2019
localization_priority: Normal
---


# ParagraphFormat.TextDirection property (Publisher)

Returns or sets a **[PbTextDirection](publisher.pbtextdirection.md)** constant indicating the direction in which text flows in the specified paragraph. Read/write.


## Syntax

_expression_.**TextDirection**

_expression_ A variable that represents a **[ParagraphFormat](Publisher.ParagraphFormat.md)** object.


## Return value

PbTextDirection


## Remarks

This property is meant to be used in conjunction with documents that have text in both left-to-right and right-to-left languages. Setting the property to a value that is not in accordance with the text direction dictated by the language in use may have unpredictable results.

The **TextDirection** property value can be one of the **PbTextDirection** constants declared in the Microsoft Publisher type library.

## Example

The following example changes the text direction of the first shape on page one so that it flows from right to left.

```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .ParagraphFormat.TextDirection = pbTextDirectionRightToLeft
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]