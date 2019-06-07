---
title: Hyperlinks.Count property (Publisher)
keywords: vbapb10.chm6881283
f1_keywords:
- vbapb10.chm6881283
ms.prod: publisher
api_name:
- Publisher.Hyperlinks.Count
ms.assetid: 36747f3e-b365-11ca-9cbe-f6148f7da235
ms.date: 06/08/2019
localization_priority: Normal
---


# Hyperlinks.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[Hyperlinks](Publisher.Hyperlinks.md)** object.


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