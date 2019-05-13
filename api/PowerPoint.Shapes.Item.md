---
title: Shapes.Item method (PowerPoint)
keywords: vbapp10.chm543003
f1_keywords:
- vbapp10.chm543003
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.Item
ms.assetid: f6c5eac1-3b65-3023-3b7a-557c7bfb0f02
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Item method (PowerPoint)

Returns a single  **Shape** object from the specified **Shapes** collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the single  **Shape** object in the collection to be returned.|

## Return value

Shape


## Example

This example sets the foreground color to red for the shape named "Rectangle 1" on slide one in the active presentation.


```vb
ActivePresentation.Slides.Item(1).Shapes.Item("rectangle 1").Fill _
    .ForeColor.RGB = RGB(128, 0, 0)
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]