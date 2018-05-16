---
title: ToggleButton.PressedThemeColorIndex Property (Access)
keywords: vbaac10.chm14620
f1_keywords:
- vbaac10.chm14620
ms.prod: access
api_name:
- Access.ToggleButton.PressedThemeColorIndex
ms.assetid: 85609290-6641-001c-7bc2-0f14443b326f
ms.date: 06/08/2017
---


# ToggleButton.PressedThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **PressedColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **PressedThemeColorIndex**

 _expression_ A variable that represents a **ToggleButton** object.


## Remarks

The  **PressedThemeColorIndex** uses one of the values listed in the following table.



|**Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1 |Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8 (Default)|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **PressedThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[ToggleButton Object](Access.ToggleButton.md)

