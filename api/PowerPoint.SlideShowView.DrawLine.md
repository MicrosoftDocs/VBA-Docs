---
title: SlideShowView.DrawLine method (PowerPoint)
keywords: vbapp10.chm513015
f1_keywords:
- vbapp10.chm513015
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.DrawLine
ms.assetid: d4c3c1c9-cd12-67ba-b1b9-4d7e924bd084
ms.date: 06/08/2017
localization_priority: Normal
---


# SlideShowView.DrawLine method (PowerPoint)

Draws a line in the specified slide show view.


## Syntax

_expression_. `DrawLine`( `_BeginX_`, `_BeginY_`, `_EndX_`, `_EndY_` )

_expression_ A variable that represents a [SlideShowView](PowerPoint.SlideShowView.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeginX_|Required|**Single**|The position (in points) of the line's starting point relative to the upper-left corner of the slide.|
| _BeginY_|Required|**Single**|The position (in points) of the line's starting point relative to the upper-left corner of the slide.|
| _EndX_|Required|**Single**|The position (in points) of the line's ending point relative to the upper-left corner of the slide.|
| _EndY_|Required|**Single**|The position (in points) of the line's ending point relative to the upper-left corner of the slide.|

## Example

This example draws a line in slide show window one.


```vb
SlideShowWindows(1).View.DrawLine 5, 5, 250, 250
```


## See also


[SlideShowView Object](PowerPoint.SlideShowView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]