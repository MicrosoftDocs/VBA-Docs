---
title: ColorStops.Add method (Excel)
keywords: vbaxl10.chm853074
f1_keywords:
- vbaxl10.chm853074
ms.prod: excel
api_name:
- Excel.ColorStops.Add
ms.assetid: 121c48bf-0b68-89c9-6a03-f7a403b52fee
ms.date: 04/20/2019
localization_priority: Normal
---


# ColorStops.Add method (Excel)

Adds a **[ColorStop](Excel.ColorStop.md)** object to the specified collection.


## Syntax

_expression_.**Add** (_Position_)

_expression_ An expression that returns a **[ColorStops](Excel.ColorStops.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Position_|Required| **Double**|Represents the position in which to apply the **ColorStop**.|

## Return value

ColorStop


## Example

Adds a **ColorStop** for the active selection.

```vb
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]