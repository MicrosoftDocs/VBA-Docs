---
title: Shapes.AddSmartArt method (PowerPoint)
keywords: vbapp10.chm543034
f1_keywords:
- vbapp10.chm543034
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddSmartArt
ms.assetid: 5bd66a76-a31c-3633-7aae-f24e0a92021c
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddSmartArt method (PowerPoint)

Adds a SmartArt diagram to the  **Shapes** object.


## Syntax

_expression_. `AddSmartArt`( `_Layout_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Layout_|Required|**[SMARTARTLAYOUT]**|The SmartArt diagram to add.|
| _Left_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the slide to the left edge of the SmartArt diagram.|
| _Top_|Optional|**Single**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the top edge of the slide to the top edge of the SmartArt diagram.|
| _Width_|Optional|**Single**|The width of the SmartArt diagram.|
| _Height_|Optional|**Single**|The height of the SmartArt diagram.|

## Return value

Shape


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]