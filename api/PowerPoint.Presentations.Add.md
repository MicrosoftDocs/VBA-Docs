---
title: Presentations.Add method (PowerPoint)
keywords: vbapp10.chm522004
f1_keywords:
- vbapp10.chm522004
ms.prod: powerpoint
api_name:
- PowerPoint.Presentations.Add
ms.assetid: 9a09ad9b-c52d-9fd6-20ef-68b694596ed2
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentations.Add method (PowerPoint)

Creates a presentation. Returns a **[Presentation](PowerPoint.Presentation.md)** object that represents the new presentation.


## Syntax

_expression_.**Add** (_WithWindow_)

_expression_ A variable that represents a [Presentations](PowerPoint.Presentations.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _WithWindow_|Optional|**MsoTriState**|Whether the presentation appears in a visible window.|

## Return value

Presentation


## Remarks

The  _WithWindow_ parameter value can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The new presentation isn't visible.|
|**msoTrue**|The default. Creates the presentation in a visible window.|

## Example

This example creates a presentation, adds a slide to it, and then saves the presentation.


```vb
With Presentations.Add

    .Slides.Add Index:=1, Layout:=ppLayoutTitle

    .SaveAs "Sample"

End With


```


## See also


[Presentations Object](PowerPoint.Presentations.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
