---
title: CatalogMergeShapes.Count property (Publisher)
keywords: vbapb10.chm8388611
f1_keywords:
- vbapb10.chm8388611
ms.prod: publisher
api_name:
- Publisher.CatalogMergeShapes.Count
ms.assetid: a871af2f-183c-f5a8-7ad0-c8d25c71e41f
ms.date: 06/06/2019
localization_priority: Normal
---


# CatalogMergeShapes.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[CatalogMergeShapes](Publisher.CatalogMergeShapes.md)** object.


## Example

This example displays the number of pages in the active document.

```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " & _ 
 ActiveDocument.Pages.Count & " page(s)." 
End Sub
```

<br/>

This example displays the number of shapes in the active document.

```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " & intShapes & " shape(s)." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]