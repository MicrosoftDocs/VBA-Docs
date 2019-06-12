---
title: MailMergeMappedDataFields.Count property (Publisher)
keywords: vbapb10.chm6488067
f1_keywords:
- vbapb10.chm6488067
ms.prod: publisher
api_name:
- Publisher.MailMergeMappedDataFields.Count
ms.assetid: 45bb34e6-3b6f-2daa-d782-2bbd02b1e7b4
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeMappedDataFields.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[MailMergeMappedDataFields](Publisher.MailMergeMappedDataFields.md)** object.


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