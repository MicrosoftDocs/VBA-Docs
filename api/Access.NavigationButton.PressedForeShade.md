---
title: NavigationButton.PressedForeShade property (Access)
keywords: vbaac10.chm14626
f1_keywords:
- vbaac10.chm14626
ms.prod: access
api_name:
- Access.NavigationButton.PressedForeShade
ms.assetid: 5a086d71-d916-5b97-9e98-51f6394f3eaa
ms.date: 03/05/2019
localization_priority: Normal
---


# NavigationButton.PressedForeShade property (Access)

Gets or sets the shade applied to the theme color in the **PressedForeColor** property of the specified object. Read/write **Single**.


## Syntax

_expression_.**PressedForeShade**

_expression_ A variable that represents a **[NavigationButton](Access.NavigationButton.md)** object.


## Remarks

The **PressedForeShade** property contains a numeric expression that can be used to darken the theme color in the **PressedForeColor** property. The default value of the **PressedForeShade** property is 100, which is neutral, and does not change the theme color. 

To darken the color, first determine the percentage by which to darken from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]