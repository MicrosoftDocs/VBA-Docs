---
title: ShadowFormat.IncrementOffsetY method (Publisher)
keywords: vbapb10.chm3670033
f1_keywords:
- vbapb10.chm3670033
ms.prod: publisher
api_name:
- Publisher.ShadowFormat.IncrementOffsetY
ms.assetid: fca7a688-adf8-d8cd-8e14-9d1988c8d9f2
ms.date: 06/13/2019
localization_priority: Normal
---


# ShadowFormat.IncrementOffsetY method (Publisher)

Incrementally changes the vertical offset of the shadow by the specified distance.


## Syntax

_expression_.**IncrementOffsetY** (_Increment_)

_expression_ A variable that represents a **[ShadowFormat](Publisher.ShadowFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Increment_|Required| **Variant**|Specifies how far the shadow offset is to be moved vertically. A positive value moves the shadow down; a negative value moves it up. Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").|

## Remarks

Use the **[OffsetY](Publisher.ShadowFormat.OffsetY.md)** property to set the absolute vertical shadow offset.

Use the **[IncrementOffsetX](Publisher.ShadowFormat.IncrementOffsetX.md)** method to change a shadow's horizontal offset.


## Example

This example moves the shadow for the third shape in the active publication up by 3 points.

```vb
ActiveDocument.Pages(1).Shapes(3).Shadow _ 
 .IncrementOffsetY Increment:=-3 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]