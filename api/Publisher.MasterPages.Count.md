---
title: MasterPages.Count property (Publisher)
keywords: vbapb10.chm589827
f1_keywords:
- vbapb10.chm589827
ms.prod: publisher
api_name:
- Publisher.MasterPages.Count
ms.assetid: adb14000-5dc4-9154-5c5f-8f63c89309b7
ms.date: 06/11/2019
localization_priority: Normal
---


# MasterPages.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[MasterPages](Publisher.MasterPages.md)** object.


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