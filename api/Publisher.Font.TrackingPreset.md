---
title: Font.TrackingPreset property (Publisher)
keywords: vbapb10.chm5373986
f1_keywords:
- vbapb10.chm5373986
ms.prod: publisher
api_name:
- Publisher.Font.TrackingPreset
ms.assetid: 818e6efd-a1b3-1ccd-1dc1-29c0a8ded7f2
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.TrackingPreset property (Publisher)

Returns or sets a **[PbTrackingPresetType](publisher.pbtrackingpresettype.md)** constant representing the preset tracking type for characters in the specified font in a text range. Read/write.


## Syntax

_expression_.**TrackingPreset**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

PbTrackingPresetType


## Remarks

The **TrackingPreset** property value can be one of the **PbTrackingPresetType** constants.

Loose and very loose tracking leaves ample space between characters, whereas tight and very tight tracking can produce character overlap.


## Example

This example specifies tight tracking as the preset for the characters in the second story.

```vb
Sub TrackingType() 
 
 Application.ActiveDocument.Stories(2).TextRange _ 
 .Font.TrackingPreset = pbTrackingTight 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]