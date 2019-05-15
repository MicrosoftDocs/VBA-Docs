---
title: Shapes.AddSmartArt method (Excel)
keywords: vbaxl10.chm638095
f1_keywords:
- vbaxl10.chm638095
ms.prod: excel
api_name:
- Excel.Shapes.AddSmartArt
ms.assetid: e18a53ef-7649-34be-a264-aa545dd3d012
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.AddSmartArt method (Excel)

Creates a new SmartArt graphic with the specified layout. 


## Syntax

_expression_.**AddSmartArt** (_Layout_, _Left_, _Top_, _Width_, _Height_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Layout_|Required| **[SmartArtLayout](Office.SmartArtLayout.md)**|An object that represents the layout to use.|
| _Left_|Optional| **Variant**|The distance, in [points](../language/glossary/vbe-glossary.md#point), from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).|
| _Top_|Optional| **Variant**|The distance, in points, from the top edge of the object to the top edge of the worksheet.|
| _Width_|Optional| **Variant**|The width, in points, of the object.|
| _Height_|Optional| **Variant**|The height, in points, of the object.|

## Return value

**Shape**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]