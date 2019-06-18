---
title: WizardProperties.Count property (Publisher)
keywords: vbapb10.chm1507331
f1_keywords:
- vbapb10.chm1507331
ms.prod: publisher
api_name:
- Publisher.WizardProperties.Count
ms.assetid: 835f3467-ec89-54d2-c685-3021e6267121
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardProperties.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[WizardProperties](Publisher.WizardProperties.md)** object.


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