---
title: GradientStops.Insert2 method (Office)
ms.prod: office
api_name:
- Office.GradientStops.Insert2
ms.assetid: bd9ed41d-eaeb-d3aa-6a8a-e38e2bfb9a17
ms.date: 01/16/2019
localization_priority: Normal
---


# GradientStops.Insert2 method (Office)

Adds a stop to a gradient, and specifies the brightness, as well as the transparency, of the color.


## Syntax

_expression_.**Insert2** (_RGB_, _Position_, _Transparency_, _Index_, _Brightness_)

_expression_ An expression that returns a **[GradientStops](Office.GradientStops.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RGB_|Required|**[MsoThemeColorSchemeIndex](office.msothemecolorschemeindex.md)**|Specifies the color at the gradient stop.|
| _Position_|Required|**Single**|Specifies the position of the stop within the gradient expressed as a percent.|
| _Transparency_|Optional|**Single**|Specifies the opacity of the color at the gradient stop.|
| _Index_|Optional|**Integer**|The index number of the gradient stop.|
| _Brightness_|Optional|**Single**|Specifies the brightness of the color at the gradient stop.|

## Return value

Nothing


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops. 

This method differs from the **[Insert](Office.GradientStops.Insert.md)** method in that it allows you to specify the brightness, as well as the transparency, of the color at the gradient stop.


## See also

- [GradientStops object members](overview/library-reference/gradientstops-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]