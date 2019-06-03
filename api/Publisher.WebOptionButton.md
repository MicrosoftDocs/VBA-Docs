---
title: WebOptionButton object (Publisher)
keywords: vbapb10.chm4325375
f1_keywords:
- vbapb10.chm4325375
ms.prod: publisher
api_name:
- Publisher.WebOptionButton
ms.assetid: acdbaebd-b333-02b1-bf4d-d7e92148a275
ms.date: 06/04/2019
localization_priority: Normal
---


# WebOptionButton object (Publisher)

Represents a web option button control. The **WebOptionButton** object is a member of the **[Shape](publisher.shape.md)** object.
 
## Remarks

Use the **[Shapes.AddWebControl](Publisher.Shapes.AddWebControl.md)** method to create a new web option button. 

Use the **[Shape.WebOptionButton](Publisher.Shape.WebOptionButton.md)** property to access a web option button control shape. 

## Example

This example creates a new web option button and specifies that its default state is selected; it then adds a text box next to it to describe it.

```vb
Sub CreateNewWebOptionButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlOptionButton, Left:=100, _ 
 Top:=123, Width:=16, Height:=10).WebOptionButton 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=120, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Advanced User" 
 End With 
 End With 
End Sub
```


## Properties

- [Application](Publisher.WebOptionButton.Application.md)
- [Parent](Publisher.WebOptionButton.Parent.md)
- [ReturnDataLabel](Publisher.WebOptionButton.ReturnDataLabel.md)
- [Selected](Publisher.WebOptionButton.Selected.md)
- [Value](Publisher.WebOptionButton.Value.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]