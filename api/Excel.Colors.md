---
title: Colors property (Excel Graph)
keywords: vbagr10.chm3082363
f1_keywords:
- vbagr10.chm3082363
api_name:
- Excel.Colors
ms.assetid: 8e848003-2ae8-a1d4-9ecf-8e6f87a5a600
ms.date: 04/10/2019
ms.localizationpriority: medium
---


# Colors property (Excel Graph)

Returns or sets colors in the palette for a **Chart** object. The palette has 56 entries, each represented by an RGB value. Read/write **Variant**.

## Syntax

_expression_.**Colors** (_Index_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_ |Optional |**Variant**| The color number (from 1 to 56). If this argument isn't specified, this method returns an array that contains all 56 of the colors in the palette.|

## Example

This example sets color five in the color palette for the active chart.

```vb
ActiveChart.Colors(5) = RGB(255, 0, 0) 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]