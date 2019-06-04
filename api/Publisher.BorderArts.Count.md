---
title: BorderArts.Count property (Publisher)
keywords: vbapb10.chm7733251
f1_keywords:
- vbapb10.chm7733251
ms.prod: publisher
api_name:
- Publisher.BorderArts.Count
ms.assetid: 024cd14d-80f7-7372-c550-ef804661bbae
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArts.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[BorderArts](Publisher.BorderArts.md)** object.


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