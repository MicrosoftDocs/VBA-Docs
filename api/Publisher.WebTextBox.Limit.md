---
title: WebTextBox.Limit property (Publisher)
keywords: vbapb10.chm4194309
f1_keywords:
- vbapb10.chm4194309
ms.prod: publisher
api_name:
- Publisher.WebTextBox.Limit
ms.assetid: b6bf334e-a610-492a-b316-e8b52d223176
ms.date: 06/18/2019
localization_priority: Normal
---


# WebTextBox.Limit property (Publisher)

Returns or sets a **Long** that represents the maximum number of characters that can be entered into a web text box control. Read/write.


## Syntax

_expression_.**Limit**

_expression_ A variable that represents a **[WebTextBox](Publisher.WebTextBox.md)** object.


## Return value

Long


## Remarks

Text box limits can be any number from 1 to 255 characters. Numbers higher than 255 will generate an error.


## Example

This example creates a new web text box control in the active publication, sets the default text and the character limit for the text box, and specifies that it is a required control.


```vb
Sub AddWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlMultiLineTextBox, Left:=72, _ 
 Top:=72, Width:=300, Height:=100).WebTextBox 
 .DefaultText = "Please enter text here." 
 .Limit = 200 
 .RequiredControl = msoTrue 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]