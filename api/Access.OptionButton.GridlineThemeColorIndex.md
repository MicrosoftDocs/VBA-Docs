---
title: OptionButton.GridlineThemeColorIndex property (Access)
keywords: vbaac10.chm14635
f1_keywords:
- vbaac10.chm14635
ms.prod: access
api_name:
- Access.OptionButton.GridlineThemeColorIndex
ms.assetid: 58a70e63-9c82-4761-8597-c134882e04e3
ms.date: 02/14/2019
localization_priority: Normal
---


# OptionButton.GridlineThemeColorIndex property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the **[GridlineColor](access.OptionButton.gridlinecolor.md)** property of the specified object. Read/write **Long**.


## Syntax

_expression_.**GridlineThemeColorIndex**

_expression_ A variable that represents an **[OptionButton](Access.OptionButton.md)** object.


## Remarks

The **GridlineThemeColorIndex** property uses one of the values listed in the following table.

|Value|Description|
|:-----|:-----|
|0 |Text 1|
|1 (Default)|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|

If no theme is applied, the **GridlineThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet. 




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]