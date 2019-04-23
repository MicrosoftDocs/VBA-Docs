---
title: CalloutFormat object (Excel)
keywords: vbaxl10.chm104000
f1_keywords:
- vbaxl10.chm104000
ms.prod: excel
api_name:
- Excel.CalloutFormat
ms.assetid: d9d7d279-04ef-dbee-23cd-ddd606ed917d
ms.date: 03/29/2019
localization_priority: Normal
---


# CalloutFormat object (Excel)

Contains properties and methods that apply to line callouts.

## Remarks

Use the **[Callout](Excel.Shape.Callout.md)** property of the **Shape** object to return a **CalloutFormat** object.


## Example

The following example specifies the attributes of shape three (a line callout) on _myDocument_: 

- The callout will have a vertical accent bar that separates the text from the callout line.
- The angle between the callout line and the side of the callout text box will be 30 degrees.
- There will be no border around the callout text.
- The callout line will be attached to the top of the callout text box.
- The callout line will contain two segments. 

For this example to work, shape three must be a callout.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Callout 
 .Accent = True 
 .Angle = msoCalloutAngle30 
 .Border = False 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
End With
```


## Methods

- [AutomaticLength](Excel.CalloutFormat.AutomaticLength.md)
- [CustomDrop](Excel.CalloutFormat.CustomDrop.md)
- [CustomLength](Excel.CalloutFormat.CustomLength.md)
- [PresetDrop](Excel.CalloutFormat.PresetDrop.md)

## Properties

- [Accent](Excel.CalloutFormat.Accent.md)
- [Angle](Excel.CalloutFormat.Angle.md)
- [Application](Excel.CalloutFormat.Application.md)
- [AutoAttach](Excel.CalloutFormat.AutoAttach.md)
- [AutoLength](Excel.CalloutFormat.AutoLength.md)
- [Border](Excel.CalloutFormat.Border.md)
- [Creator](Excel.CalloutFormat.Creator.md)
- [Drop](Excel.CalloutFormat.Drop.md)
- [DropType](Excel.CalloutFormat.DropType.md)
- [Gap](Excel.CalloutFormat.Gap.md)
- [Length](Excel.CalloutFormat.Length.md)
- [Parent](Excel.CalloutFormat.Parent.md)
- [Type](Excel.CalloutFormat.Type.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]