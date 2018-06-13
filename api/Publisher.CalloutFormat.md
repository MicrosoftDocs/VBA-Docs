---
title: CalloutFormat Object (Publisher)
keywords: vbapb10.chm2555903
f1_keywords:
- vbapb10.chm2555903
ms.prod: publisher
api_name:
- Publisher.CalloutFormat
ms.assetid: 1f54aba3-3872-e668-fe76-1966d1a62cca
ms.date: 06/08/2017
---


# CalloutFormat Object (Publisher)

Contains properties and methods that apply to line callouts.
 


## Example

Use the  **[Callout](Publisher.Shape.Callout.md)** property to return a **CalloutFormat** object. The following example adds a callout to the active publication, adds text to the callout, then specifies the following attributes for the callout:
 

 

 

 

- a vertical accent bar that separates the text from the callout line ( **Accent** property)
    
 
- the angle between the callout line and the side of the callout text box will be 30 degrees ( **Angle** property)
    
 
- there will be no border around the callout text ( **Border** property)
    
 
- the callout line will be attached to the top of the callout text box ( **PresetDrop** method)
    
 
- the callout line will contain three segments ( **Type** property)
    
 



```
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



|**Name**|
|:-----|
|[AutomaticLength](Publisher.CalloutFormat.AutomaticLength.md)|
|[CustomDrop](Publisher.CalloutFormat.CustomDrop.md)|
|[CustomLength](Publisher.CalloutFormat.CustomLength.md)|
|[PresetDrop](Publisher.CalloutFormat.PresetDrop.md)|

## Properties



|**Name**|
|:-----|
|[Accent](Publisher.CalloutFormat.Accent.md)|
|[Angle](Publisher.CalloutFormat.Angle.md)|
|[Application](Publisher.CalloutFormat.Application.md)|
|[AutoAttach](Publisher.CalloutFormat.AutoAttach.md)|
|[AutoLength](Publisher.CalloutFormat.AutoLength.md)|
|[Border](Publisher.CalloutFormat.Border.md)|
|[Drop](Publisher.CalloutFormat.Drop.md)|
|[DropType](Publisher.CalloutFormat.DropType.md)|
|[Gap](Publisher.CalloutFormat.Gap.md)|
|[Length](Publisher.CalloutFormat.Length.md)|
|[Parent](Publisher.CalloutFormat.Parent.md)|
|[Type](Publisher.CalloutFormat.Type.md)|

