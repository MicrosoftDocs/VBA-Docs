---
title: FillFormat.GradientColorType Property (Publisher)
keywords: vbapb10.chm2359554
f1_keywords:
- vbapb10.chm2359554
ms.prod: publisher
api_name:
- Publisher.FillFormat.GradientColorType
ms.assetid: b0866675-4bc4-5e82-780d-8bae4b7d16ef
ms.date: 06/08/2017
---


# FillFormat.GradientColorType Property (Publisher)

Returns an  **MsoGradientColorType** constant indicating the gradient color type for the specified fill. Read-only.


## Syntax

 _expression_. **GradientColorType**

 _expression_ A variable that represents a  **FillFormat** object.


### Return Value

MsoGradientColorType


## Remarks

Use the  [OneColorGradient](Publisher.FillFormat.OneColorGradient.md),  [PresetGradient](Publisher.FillFormat.PresetGradient.md), or  **[TwoColorGradient](Publisher.FillFormat.TwoColorGradient.md)** method to set the gradient type for the fill.

The  **GradientColorType** property value can be one of the ** [MsoGradientColorType](./Office.MsoGradientColorType.md)** constants declared in the Microsoft Office type library.


## Example

This example changes the fill for all shapes on the first page of the active publication that have a two-color gradient fill to a preset gradient fill.


```vb
Dim shpLoop As Shape 
 
' Loop through collection of shapes. 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop.Fill 
 ' Test for two-color gradient. 
 If .GradientColorType = msoGradientTwoColors Then 
 ' Apply a preset gradient. 
 .PresetGradient Style:=msoGradientHorizontal, _ 
 Variant:=1, PresetGradientType:=msoGradientBrass 
 End If 
 End With 
Next shpLoop 

```


