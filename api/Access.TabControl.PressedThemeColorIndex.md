---
title: TabControl.PressedThemeColorIndex property (Access)
keywords: vbaac10.chm14620
f1_keywords:
- vbaac10.chm14620
ms.prod: access
api_name:
- Access.TabControl.PressedThemeColorIndex
ms.assetid: 1ac09ba3-ba29-81d1-044a-210cd09e1cb6
ms.date: 06/08/2017
---


# TabControl.PressedThemeColorIndex property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **PressedColor** property of the specified object. Read/write **Long**.


## Syntax

_expression_. `PressedThemeColorIndex`

_expression_ A variable that represents a [TabControl](Access.TabControl.md) object.


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


[TabControl Object](Access.TabControl.md)

