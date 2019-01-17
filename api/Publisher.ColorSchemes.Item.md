---
title: ColorSchemes.Item Property (Publisher)
keywords: vbapb10.chm2752512
f1_keywords:
- vbapb10.chm2752512
ms.prod: publisher
api_name:
- Publisher.ColorSchemes.Item
ms.assetid: 5a66a0ae-b552-0979-d3ac-7b1d7bec96f7
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorSchemes.Item Property (Publisher)

Returns the specified  **[ColorScheme](Publisher.ColorScheme.md)** object from a **ColorSchemes** collection. Read-only.


## Syntax

 _expression_. **Item**(**_Index_**)

 _expression_ A variable that represents a  **ColorSchemes** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The color scheme to return. Can be either the name of the color scheme as a string or the corresponding  **PbColorScheme** constant.|

## Remarks

The  **Item** property value can be one of the **[PbColorScheme](Publisher.PbColorScheme.md)** constants declared in the Microsoft Publisher type library.


## Example

This example sets the color scheme of the active publication to the Aqua color scheme.


```vb
ActiveDocument.ColorScheme = ColorSchemes.Item(Index:=pbColorSchemeAqua)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]