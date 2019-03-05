---
title: NavigationButton.HoverForeThemeColorIndex property (Access)
keywords: vbaac10.chm14616
f1_keywords:
- vbaac10.chm14616
ms.prod: access
api_name:
- Access.NavigationButton.HoverForeThemeColorIndex
ms.assetid: 0fe67489-953c-294b-a226-c746e0321782
ms.date: 03/05/2019
localization_priority: Normal
---


# NavigationButton.HoverForeThemeColorIndex property (Access)

Gets or sets the theme color index that represents a color in the applied color theme associated with the **HoverForeColor** property of the specified object. Read/write **Long**.


## Syntax

_expression_.**HoverForeThemeColorIndex**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

The **HoverForeThemeColorIndex** property uses one of the values listed in the following table.

|Value|Description|
|:-----|:-----|
|0|Text 1|
|1|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7 (Default)|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|

If no theme is applied, the **HoverForeThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]