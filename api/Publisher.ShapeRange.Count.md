---
title: ShapeRange.Count property (Publisher)
keywords: vbapb10.chm2293763
f1_keywords:
- vbapb10.chm2293763
ms.prod: publisher
api_name:
- Publisher.ShapeRange.Count
ms.assetid: 5037bfe9-b430-4205-c514-b2f4313b4c53
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


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