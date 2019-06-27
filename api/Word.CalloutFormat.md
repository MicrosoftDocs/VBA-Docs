---
title: CalloutFormat object (Word)
keywords: vbawd10.chm2501
f1_keywords:
- vbawd10.chm2501
ms.prod: word
api_name:
- Word.CalloutFormat
ms.assetid: d54764e6-d761-582b-aa0a-baebd3a7cf6a
ms.date: 06/08/2017
localization_priority: Normal
---


# CalloutFormat object (Word)

Contains properties and methods that apply to line callouts.


## Remarks

Use the  **Callout** property to return a **CalloutFormat** object. The following example specifies the following attributes of shape three (a line callout) on the active document: the callout will have a vertical accent bar that separates the text from the callout line; the angle between the callout line and the side of the callout text box will be 30 degrees; there will be no border around the callout text; the callout line will be attached to the top of the callout text box; and the callout line will contain two segments. For this example to work, shape three must be a callout.


```vb
With ActiveDocument.Shapes(3).Callout 
 .Accent = True 
 .Angle = msoCalloutAngle30 
 .Border = False 
 .PresetDrop msoCalloutDropTop 
 .Type = msoCalloutThree 
End With
```

## Methods

- [CustomDrop](Word.CalloutFormat.CustomDrop.md)
- [CustomLength](Word.CalloutFormat.CustomLength.md)
- [PresetDrop](Word.CalloutFormat.PresetDrop.md)

## Properties

- [Accent](Word.CalloutFormat.Accent.md)
- [Angle](Word.CalloutFormat.Angle.md)
- [Application](Word.CalloutFormat.Application.md)
- [AutoLength](Word.CalloutFormat.AutoLength.md)
- [Border](Word.CalloutFormat.Border.md)
- [Creator](Word.CalloutFormat.Creator.md)
- [Drop](Word.CalloutFormat.Drop.md)
- [DropType](Word.CalloutFormat.DropType.md)
- [Gap](Word.CalloutFormat.Gap.md)
- [Length](Word.CalloutFormat.Length.md)
- [Parent](Word.CalloutFormat.Parent.md)
- [Type](Word.CalloutFormat.Type.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]