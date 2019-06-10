---
title: ObjectVerbs.Count property (Publisher)
keywords: vbapb10.chm4521987
f1_keywords:
- vbapb10.chm4521987
ms.prod: publisher
api_name:
- Publisher.ObjectVerbs.Count
ms.assetid: 0d868be0-f46d-d8bb-2af1-47e2d1a3a388
ms.date: 06/11/2019
localization_priority: Normal
---


# ObjectVerbs.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an **[ObjectVerbs](Publisher.ObjectVerbs.md)** object.


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