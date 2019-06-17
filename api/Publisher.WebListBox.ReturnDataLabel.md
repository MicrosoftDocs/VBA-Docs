---
title: WebListBox.ReturnDataLabel property (Publisher)
keywords: vbapb10.chm4063237
f1_keywords:
- vbapb10.chm4063237
ms.prod: publisher
api_name:
- Publisher.WebListBox.ReturnDataLabel
ms.assetid: 0c9a6942-1cc7-92b6-116e-836e79560084
ms.date: 06/18/2019
localization_priority: Normal
---


# WebListBox.ReturnDataLabel property (Publisher)

Returns or sets a **String** that represents the text used by the webpage to label the specified web object when the page is submitted. Read/write.


## Syntax

_expression_.**ReturnDataLabel**

_expression_ A variable that represents a **[WebListBox](Publisher.WebListBox.md)** object.


## Example

This example creates a new web text box and specifies the label for the text in the text box when the page is submitted.

```vb
Sub LabelWebTextBoxControl() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=300, Height:=15).WebTextBox 
 .DefaultText = "Please enter your name here" 
 .Limit = 70 
 .RequiredControl = msoTrue 
 .ReturnDataLabel = "Full_Name" 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]