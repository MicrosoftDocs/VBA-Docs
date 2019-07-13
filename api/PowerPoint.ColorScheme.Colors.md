---
title: ColorScheme.Colors method (PowerPoint)
keywords: vbapp10.chm537003
f1_keywords:
- vbapp10.chm537003
ms.prod: powerpoint
api_name:
- PowerPoint.ColorScheme.Colors
ms.assetid: ac910a40-9014-e709-491c-a8649fc08137
ms.date: 06/08/2017
localization_priority: Normal
---


# ColorScheme.Colors method (PowerPoint)

Returns an **[RGBColor](PowerPoint.RGBColor.md)** object that represents a single color in a color scheme.


## Syntax

_expression_.**Colors** (_SchemeColor_)

_expression_ A variable that represents a **[ColorScheme](PowerPoint.ColorScheme.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SchemeColor_|Required|**[PpColorSchemeIndex](PowerPoint.PpColorSchemeIndex.md)**|The individual color in the specified color scheme.|

## Return value

RGBColor


## Example

This example sets the title color for slides one and three in the active presentation.


```vb
Set mySlides = ActivePresentation.Slides.Range(Array(1, 3))

mySlides.ColorScheme.Colors(ppTitle).RGB = RGB(0, 255, 0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]