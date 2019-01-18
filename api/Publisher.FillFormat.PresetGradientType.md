---
title: FillFormat.PresetGradientType Property (Publisher)
keywords: vbapb10.chm2359559
f1_keywords:
- vbapb10.chm2359559
ms.prod: publisher
api_name:
- Publisher.FillFormat.PresetGradientType
ms.assetid: 1febce3f-9925-3fec-eb78-f5167585477e
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.PresetGradientType Property (Publisher)

Returns an  **MsoPresetGradientType** constant that represents the preset gradient type for the specified fill. Read-only.


## Syntax

 _expression_. **PresetGradientType**

 _expression_ A variable that represents a  **FillFormat** object.


## Return value

MsoPresetGradientType


## Remarks

The  **PresetGradientType** property value can be one of the ** [MsoPresetGradientType](Office.MsoPresetGradientType.md)** constants declared in the Microsoft Office type library.

Use the  **[PresetGradient](Publisher.FillFormat.PresetGradient.md)** method to set the preset gradient type for the fill.


## Example

This example changes the fill for the first shape on the first page of the active publication to the Fog preset gradient fill if it is not already set to the Fog preset gradient. This example assumes that there is at least one shape on the first page of the active publication.


```vb
Sub SetGradient() 
 With ActiveDocument.Pages(1).Shapes(1).Fill 
 If .PresetGradientType <> msoGradientFog Then 
 .PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientFog 
 End If 
 End With 
End Sub
```


