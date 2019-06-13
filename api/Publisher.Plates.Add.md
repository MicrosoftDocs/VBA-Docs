---
title: Plates.Add method (Publisher)
keywords: vbapb10.chm2818052
f1_keywords:
- vbapb10.chm2818052
ms.prod: publisher
api_name:
- Publisher.Plates.Add
ms.assetid: 7fb7b602-8797-e275-4ff7-2e87cf1db11f
ms.date: 06/13/2019
localization_priority: Normal
---


# Plates.Add method (Publisher)

Adds a new color plate to the specified **Plates** object.


## Syntax

_expression_.**Add** (_PlateColor_)

_expression_ A variable that represents a **[Plates](Publisher.Plates.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PlateColor_|Optional| **[ColorFormat](publisher.colorformat.md)**| The color settings to apply to the new plate.|

## Remarks

If the **ColorMode** property of the specified publication is not **pbColorModeSpot** or **pbColorModeSpotAndProcess**, an error occurs.


## Example

The following example adds a color plate to the active publication if it is a spot-color publication.

```vb
If ActiveDocument.ColorMode = pbColorModeSpot Then 
 ActiveDocument.Plates.Add 
End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]