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


## See also


[Word Object Model Reference](./overview/Word/object-model.md)


