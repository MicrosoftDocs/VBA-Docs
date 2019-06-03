---
title: Adjustments.Count property (Publisher)
keywords: vbapb10.chm2424835
f1_keywords:
- vbapb10.chm2424835
ms.prod: publisher
api_name:
- Publisher.Adjustments.Count
ms.assetid: 1b32f1c3-0bbc-a175-4f59-36cc76df12fd
ms.date: 06/04/2019
localization_priority: Normal
---


# Adjustments.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents an **[Adjustments](Publisher.Adjustments.md)** object.


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