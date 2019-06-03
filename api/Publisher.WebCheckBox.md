---
title: WebCheckBox object (Publisher)
keywords: vbapb10.chm4390911
f1_keywords:
- vbapb10.chm4390911
ms.prod: publisher
api_name:
- Publisher.WebCheckBox
ms.assetid: adcdf233-50b8-acbe-e52f-1e86e175b31d
ms.date: 06/04/2019
localization_priority: Normal
---


# WebCheckBox object (Publisher)

Represents a web check box control. The **WebCheckBox** object is a member of the **[Shape](publisher.shape.md)** object.
 
## Remarks

Use the **[Shapes.AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create a web check box. 

Use the **[Shape.WebCheckBox](Publisher.Shape.WebCheckBox.md)** property to access a web check box control shape. 

## Example

This example creates a new web check box and specifies that its default state is selected; it then adds a text box next to it to describe it.

```vb
Sub CreateNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12).WebCheckBox 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=118, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Description text for web check box" 
 End With 
 End With 
End Sub
```


## Properties

- [Application](Publisher.WebCheckBox.Application.md)
- [Parent](Publisher.WebCheckBox.Parent.md)
- [ReturnDataLabel](Publisher.WebCheckBox.ReturnDataLabel.md)
- [Selected](Publisher.WebCheckBox.Selected.md)
- [Value](Publisher.WebCheckBox.Value.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]