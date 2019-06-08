---
title: InlineShapes.Count property (Publisher)
keywords: vbapb10.chm5767171
f1_keywords:
- vbapb10.chm5767171
ms.prod: publisher
api_name:
- Publisher.InlineShapes.Count
ms.assetid: a78ae487-e7d6-1099-236f-6464c601686f
ms.date: 06/08/2019
localization_priority: Normal
---


# InlineShapes.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an **[InlineShapes](Publisher.InlineShapes.md)** object.


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