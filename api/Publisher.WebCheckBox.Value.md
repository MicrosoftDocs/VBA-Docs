---
title: WebCheckBox.Value property (Publisher)
keywords: vbapb10.chm4325381
f1_keywords:
- vbapb10.chm4325381
ms.prod: publisher
api_name:
- Publisher.WebCheckBox.Value
ms.assetid: 9fd50cd5-ecf3-30b7-c8a9-6b64b106eaec
ms.date: 06/15/2019
localization_priority: Normal
---


# WebCheckBox.Value property (Publisher)

Returns or sets a **String** that represents the value of a web check box or option button. Read/write.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a **[WebCheckBox](Publisher.WebCheckBox.md)** object.


## Example

This example creates a new web check box control, assigns a name and value to it, and indicates that its initial state is selected.

```vb
Sub CreateWebButton() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlCheckBox, Left:=72, _ 
 Top:=72, Width:=100, Height:=50) 
 .Name = "ControlBox" 
 With .WebCheckBox 
 .Value = "This is a check box." 
 .Selected = msoTrue 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]