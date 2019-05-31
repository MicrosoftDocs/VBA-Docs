---
title: CalloutFormat object (Publisher)
keywords: vbapb10.chm2555903
f1_keywords:
- vbapb10.chm2555903
ms.prod: publisher
api_name:
- Publisher.CalloutFormat
ms.assetid: 1f54aba3-3872-e668-fe76-1966d1a62cca
ms.date: 05/31/2019
localization_priority: Normal
---


# CalloutFormat object (Publisher)

Contains properties and methods that apply to line callouts.

## Remarks

Use the **[Callout](Publisher.Shape.Callout.md)** property of the **Shape** object to return a **CalloutFormat** object. 

## Example

The following example adds a callout to the active publication, adds text to the callout, and then specifies the following attributes for the callout:

- A vertical accent bar separates the text from the callout line (**Accent** property).
- The angle between the callout line and the side of the callout text box is 30 degrees (**Angle** property).
- There is no border around the callout text (**Border** property).
- The callout line is attached to the top of the callout text box (**PresetDrop** method).
- The callout line contains three segments (**Type** property).
    

```vb
Sub AddFormatCallout() 
 With ActiveDocument.Pages(1).Shapes.AddCallout(Type:=msoCalloutOne, _ 
 Left:=150, Top:=150, Width:=200, Height:=100) 
 With .TextFrame.TextRange 
 .Text = "This is a callout." 
 With .Font 
 .Name = "Stencil" 
 .Bold = msoTrue 
 .Size = 30 
 End With 
 End With 
 With .Callout 
 .Accent = MsoTrue 
 .Angle = msoCalloutAngle30 
 .Border = MsoFalse 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
 End With 
 End With 
End Sub
```


## Methods

- [AutomaticLength](Publisher.CalloutFormat.AutomaticLength.md)
- [CustomDrop](Publisher.CalloutFormat.CustomDrop.md)
- [CustomLength](Publisher.CalloutFormat.CustomLength.md)
- [PresetDrop](Publisher.CalloutFormat.PresetDrop.md)

## Properties

- [Accent](Publisher.CalloutFormat.Accent.md)
- [Angle](Publisher.CalloutFormat.Angle.md)
- [Application](Publisher.CalloutFormat.Application.md)
- [AutoAttach](Publisher.CalloutFormat.AutoAttach.md)
- [AutoLength](Publisher.CalloutFormat.AutoLength.md)
- [Border](Publisher.CalloutFormat.Border.md)
- [Drop](Publisher.CalloutFormat.Drop.md)
- [DropType](Publisher.CalloutFormat.DropType.md)
- [Gap](Publisher.CalloutFormat.Gap.md)
- [Length](Publisher.CalloutFormat.Length.md)
- [Parent](Publisher.CalloutFormat.Parent.md)
- [Type](Publisher.CalloutFormat.Type.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]