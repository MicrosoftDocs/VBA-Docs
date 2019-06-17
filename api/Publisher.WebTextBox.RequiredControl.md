---
title: WebTextBox.RequiredControl property (Publisher)
keywords: vbapb10.chm4194310
f1_keywords:
- vbapb10.chm4194310
ms.prod: publisher
api_name:
- Publisher.WebTextBox.RequiredControl
ms.assetid: 32e18d4b-7af0-b079-4baf-9acc07c3c37d
ms.date: 06/18/2019
localization_priority: Normal
---


# WebTextBox.RequiredControl property (Publisher)

Specifies whether an entry into a web text box control is required. Read/write.


## Syntax

_expression_.**RequiredControl**

_expression_ A variable that represents a **[WebTextBox](Publisher.WebTextBox.md)** object.


## Return value

MsoTriState


## Remarks

The **RequiredControl** property value can be one of the **[MsoTriState](office.msotristate.md)** constants declared in the Microsoft Office type library and shown in the following table.

|Constant|Description|
|:-----|:-----|
| **msoFalse**|Indicates that entry into the specified web text box control is not required.|
| **msoTrue** |Indicates that entry into the specified web text box control is required.|

## Example

This example creates a new web text box control in the active publication, sets the default text and the character limit for the text box, and specifies that an entry is required.

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