---
title: WebCheckBox Object (Publisher)
keywords: vbapb10.chm4390911
f1_keywords:
- vbapb10.chm4390911
ms.prod: publisher
api_name:
- Publisher.WebCheckBox
ms.assetid: adcdf233-50b8-acbe-e52f-1e86e175b31d
ms.date: 06/08/2017
localization_priority: Normal
---


# WebCheckBox Object (Publisher)

Represents a Web check box control. The  **WebCheckBox** object is a member of the **Shape** object.
 


## Example

Use the  **[AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create a Web check box. Use the **[WebCheckBox](Publisher.Shape.WebCheckBox.md)** property to access a Web check box control shape. This example creates a new Web check box and specifies that its default state is checked; then it adds a text box next to it to describe it.
 

 

```vb
Sub CreateNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12).WebCheckBox 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=118, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Description text for Web check box" 
 End With 
 End With 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Publisher.WebCheckBox.Application.md)|
|[Parent](Publisher.WebCheckBox.Parent.md)|
|[ReturnDataLabel](Publisher.WebCheckBox.ReturnDataLabel.md)|
|[Selected](Publisher.WebCheckBox.Selected.md)|
|[Value](Publisher.WebCheckBox.Value.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]