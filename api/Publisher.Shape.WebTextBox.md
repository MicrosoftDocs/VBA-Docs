---
title: Shape.WebTextBox property (Publisher)
keywords: vbapb10.chm2228342
f1_keywords:
- vbapb10.chm2228342
ms.prod: publisher
api_name:
- Publisher.Shape.WebTextBox
ms.assetid: 8a3f8389-728f-b8ae-3c89-dc8d03a3818e
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.WebTextBox property (Publisher)

Returns the **[WebTextBox](Publisher.WebTextBox.md)** object associated with the specified shape.


## Syntax

_expression_.**WebTextBox**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Return value

WebTextBox


## Example

This example creates a new web text box, specifies default text, indicates that entry is required, and limits entry to 50 characters.

```vb
Dim shpNew As Shape 
Dim wtbTemp As WebTextBox 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=150, Height:=15) 
 
Set wtbTemp = shpNew.WebTextBox 
 
With wtbTemp 
.DefaultText = "Please Enter Your Full Name" 
 .RequiredControl = msoTrue 
 .Limit = 50 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]