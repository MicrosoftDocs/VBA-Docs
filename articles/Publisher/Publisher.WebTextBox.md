---
title: WebTextBox Object (Publisher)
keywords: vbapb10.chm4259839
f1_keywords:
- vbapb10.chm4259839
ms.prod: publisher
api_name:
- Publisher.WebTextBox
ms.assetid: 74fde391-734c-6672-dadb-59bc58232c0f
ms.date: 06/08/2017
---


# WebTextBox Object (Publisher)

Represents a Web text box control. The  **WebTextBox** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create new Web option button. Use the **[WebTextBox](Publisher.Shape.WebTextBox.md)** property to access a Web text box control shape. This example creates a new Web text box, specifies default text, indicates that entry is required, and limits entry to 50 characters.
 

 

```
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



|**Name**|
|:-----|
|[Application](Publisher.WebTextBox.Application.md)|
|[DefaultText](Publisher.WebTextBox.DefaultText.md)|
|[EchoAsterisks](Publisher.WebTextBox.EchoAsterisks.md)|
|[Limit](Publisher.WebTextBox.Limit.md)|
|[Parent](Publisher.WebTextBox.Parent.md)|
|[RequiredControl](Publisher.WebTextBox.RequiredControl.md)|
|[ReturnDataLabel](Publisher.WebTextBox.ReturnDataLabel.md)|

