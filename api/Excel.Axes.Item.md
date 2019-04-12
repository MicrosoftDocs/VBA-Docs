---
title: Axes.Item method (Excel)
keywords: vbaxl10.chm572074
f1_keywords:
- vbaxl10.chm572074
ms.prod: excel
api_name:
- Excel.Axes.Item
ms.assetid: 5e89a576-d2a0-d069-4db6-fc1cf9bd6c61
ms.date: 04/13/2019
localization_priority: Normal
---


# Axes.Item method (Excel)

Returns a single **[Axis](Excel.Axis(object).md)** object from an **Axes** collection.


## Syntax

_expression_.**Item** (_Type_, _AxisGroup_)

_expression_ A variable that represents an **[Axes](Excel.Axes(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlAxisType](Excel.XlAxisType.md)**|The axis type.|
| _AxisGroup_|Optional| **[XlAxisGroup](Excel.XlAxisGroup.md)**|The axis.|

## Return value

Axis


## Example

This example sets the title text for the category axis on Chart1.

```vb
With Charts("chart1").Axes.Item(xlCategory) 
 .HasTitle = True 
 .AxisTitle.Caption = "1994" 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]