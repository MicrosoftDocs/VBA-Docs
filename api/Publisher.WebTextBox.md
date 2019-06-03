---
title: WebTextBox object (Publisher)
keywords: vbapb10.chm4259839
f1_keywords:
- vbapb10.chm4259839
ms.prod: publisher
api_name:
- Publisher.WebTextBox
ms.assetid: 74fde391-734c-6672-dadb-59bc58232c0f
ms.date: 06/04/2019
localization_priority: Normal
---


# WebTextBox object (Publisher)

Represents a web text box control. The **WebTextBox** object is a member of the **[Shape](publisher.shape.md)** object.
 

## Remarks

Use the **[Shapes.AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create a new web text box. 

Use the **[Shape.WebTextBox](Publisher.Shape.WebTextBox.md)** property to access a web text box control shape. 

## Example

This example creates a new web text box, specifies default text, indicates that entry is required, and limits entry to 50 characters.

```vb
Sub CreateWebTextBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlSingleLineTextBox, _ 
 Left:=100, Top:=100, Width:=150, Height:=15).WebTextBox 
 .DefaultText = "Please Enter Your Full Name" 
 .RequiredControl = msoTrue 
 .Limit = 50 
 End With 
 End With 
End Sub
```


## Properties

- [Application](Publisher.WebTextBox.Application.md)
- [DefaultText](Publisher.WebTextBox.DefaultText.md)
- [EchoAsterisks](Publisher.WebTextBox.EchoAsterisks.md)
- [Limit](Publisher.WebTextBox.Limit.md)
- [Parent](Publisher.WebTextBox.Parent.md)
- [RequiredControl](Publisher.WebTextBox.RequiredControl.md)
- [ReturnDataLabel](Publisher.WebTextBox.ReturnDataLabel.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]