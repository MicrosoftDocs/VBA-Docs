---
title: Range.BorderAround method (Excel)
keywords: vbaxl10.chm144252
f1_keywords:
- vbaxl10.chm144252
ms.prod: excel
api_name:
- Excel.Range.BorderAround
ms.assetid: 3ffeb131-45f7-7799-e04a-11577fedaa16
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.BorderAround method (Excel)

Adds a border to a range and sets the **[Color](Excel.Border.Color.md)**, **[LineStyle](Excel.Border.LineStyle.md)**, and **[Weight](Excel.Border.Weight.md)** properties of the **Border** object for the new border. **Variant**.


## Syntax

_expression_.**BorderAround** (_LineStyle_, _Weight_, _ColorIndex_, _Color_, _ThemeColor_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _LineStyle_|Optional| **Variant**|One of the constants of **[XlLineStyle](Excel.XlLineStyle.md)** specifying the line style for the border.|
| _Weight_|Optional| **[XlBorderWeight](Excel.XlBorderWeight.md)**|The border weight.|
| _ColorIndex_|Optional| **[XlColorIndex](Excel.XlColorIndex.md)**|The border color, as an index into the current color palette or as an **XlColorIndex** constant.|
| _Color_|Optional| **Variant**|The border color, as an RGB value.|
| _ThemeColor_|Optional| **Variant**|The theme color, as an index into the current color theme or as an **[XlThemeColor](Excel.XlThemeColor.md)** value.|

## Return value

Variant


## Remarks

You must specify only one of the following: _ColorIndex_, _Color_, or _ThemeColor_.

You can specify either _LineStyle_ or _Weight_, but not both. If you don't specify either argument, Microsoft Excel uses the default line style and weight.

This method outlines the entire range without filling it in. To set the borders of all the cells, you must set the **Color**, **LineStyle**, and **Weight** properties for the **[Borders](Excel.Borders.md)** collection. To clear the border, you must set the **LineStyle** property to **xlLineStyleNone** for all the cells in the range.


## Example

This example adds a thick red border around the range A1:D4 on Sheet1.

```vb
Worksheets("Sheet1").Range("A1:D4").BorderAround _ 
 ColorIndex:=3, Weight:=xlThick
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
