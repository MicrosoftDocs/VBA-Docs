---
title: ColorScheme.Colors property (Publisher)
keywords: vbapb10.chm2686978
f1_keywords:
- vbapb10.chm2686978
ms.prod: publisher
api_name:
- Publisher.ColorScheme.Colors
ms.assetid: e6599096-3f99-e7ca-0c38-1cc7d4e0a1cd
ms.date: 06/06/2019
localization_priority: Normal
---


# ColorScheme.Colors property (Publisher)

Returns a **[ColorFormat](Publisher.ColorFormat.md)** object representing a color from the specified color scheme.


## Syntax

_expression_.**Colors** (_ColorIndex_)

_expression_ A variable that represents a **[ColorScheme](Publisher.ColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ColorIndex_|Required| **[PbSchemeColorIndex](Publisher.PbSchemeColorIndex.md)**| The color from the scheme to return based on its function in the scheme. Can be one of the **PbSchemeColorIndex** constants declared in the Microsoft Publisher type library.|

## Return value

ColorFormat


## Example

The following example loops through the **[ColorSchemes](Publisher.ColorSchemes.md)** collection and looks for color schemes where the followed hyperlink color matches the color with the RGB value of 128.

```vb
Dim cscLoop As ColorScheme 
Dim colTemp As ColorFormat 
 
For Each cscLoop In Application.ColorSchemes 
 With cscLoop 
 Set colTemp = .Colors(ColorIndex:=pbSchemeColorFollowedHyperlink) 
 If colTemp.RGB = RGB(128, 0, 0) Then 
 Debug.Print "Color scheme '" & .Name _ 
 & "' has a followed hyperlink " _ 
 & "color matching RGB(128, 0, 0)" 
 End If 
 End With 
Next cscLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]