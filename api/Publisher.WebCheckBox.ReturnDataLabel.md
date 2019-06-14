---
title: WebCheckBox.ReturnDataLabel property (Publisher)
keywords: vbapb10.chm4325379
f1_keywords:
- vbapb10.chm4325379
ms.prod: publisher
api_name:
- Publisher.WebCheckBox.ReturnDataLabel
ms.assetid: a6bdc95c-d9bf-5d64-aacc-cb2e5432f3f0
ms.date: 06/15/2019
localization_priority: Normal
---


# WebCheckBox.ReturnDataLabel property (Publisher)

Returns or sets a **String** that represents the text used by the webpage to label the specified web object when the page is submitted. Read/write.


## Syntax

_expression_.**ReturnDataLabel**

_expression_ A variable that represents a **[WebCheckBox](Publisher.WebCheckBox.md)** object.


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