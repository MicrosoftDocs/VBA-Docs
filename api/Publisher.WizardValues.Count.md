---
title: WizardValues.Count property (Publisher)
keywords: vbapb10.chm1638403
f1_keywords:
- vbapb10.chm1638403
ms.prod: publisher
api_name:
- Publisher.WizardValues.Count
ms.assetid: f32f3e88-fe3e-6d47-3579-c017e4fa2994
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardValues.Count property (Publisher)

Returns a **Long** that represents the number of items in the specified collection.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a **[WizardValues](Publisher.WizardValues.md)** object.


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