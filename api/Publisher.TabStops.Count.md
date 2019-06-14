---
title: TabStops.Count property (Publisher)
keywords: vbapb10.chm5570563
f1_keywords:
- vbapb10.chm5570563
ms.prod: publisher
api_name:
- Publisher.TabStops.Count
ms.assetid: 5ba876e2-b1c0-4de9-6942-02e6688aa169
ms.date: 06/15/2019
localization_priority: Normal
---


# TabStops.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[TabStops](Publisher.TabStops.md)** object.


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