---
title: CommandButton.PressedForeThemeColorIndex Property (Access)
keywords: vbaac10.chm14624
f1_keywords:
- vbaac10.chm14624
ms.prod: access
api_name:
- Access.CommandButton.PressedForeThemeColorIndex
ms.assetid: 32ad73cd-3960-1516-c45d-175c7d642847
ms.date: 06/08/2017
---


# CommandButton.PressedForeThemeColorIndex Property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the  **PressedForeColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **PressedForeThemeColorIndex**

 _expression_ A variable that represents a **CommandButton** object.


## Remarks

The  **PressedForeThemeColorIndex** uses one of the values listed in the following table.



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
If no theme is applied, the  **PressedForeThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## See also


#### Concepts


[CommandButton Object](Access.CommandButton.md)

