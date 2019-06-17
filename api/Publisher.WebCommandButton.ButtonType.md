---
title: WebCommandButton.ButtonType property (Publisher)
keywords: vbapb10.chm3932178
f1_keywords:
- vbapb10.chm3932178
ms.prod: publisher
api_name:
- Publisher.WebCommandButton.ButtonType
ms.assetid: 9ccec0bc-4f0a-9851-0066-05ee1f144c5c
ms.date: 06/18/2019
localization_priority: Normal
---


# WebCommandButton.ButtonType property (Publisher)

Returns or sets a **[PbCommandButtonType](Publisher.PbCommandButtonType.md)** constant that indicates whether a web command button clears or submits form data. Read/write.


## Syntax

_expression_.**ButtonType**

_expression_ A variable that represents a **[WebCommandButton](Publisher.WebCommandButton.md)** object.


## Return value

PbCommandButtonType


## Remarks

The **ButtonType** property value can be one of the **PbCommandButtonType** constants declared in the Microsoft Publisher type library.


## Example

This example creates a new web command button, assigns text to appear on its face, and specifies an email address to which to send the form data.

```vb
Sub NewWebForm() 
 With ActiveDocument.Pages.Add(Count:=1, After:=1) 
 With .Shapes.AddWebControl(Type:=pbWebControlCommandButton, _ 
 Left:=72, Top:=72, Width:=72, Height:=36) 
 With .WebCommandButton 
 .ButtonType = pbCommandButtonSubmit 
 .ButtonText = "Send Form:" 
 .EmailAddress = "someone@example.com" 
 End With 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]