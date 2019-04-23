---
title: Section.AlternateBackShade property (Access)
keywords: vbaac10.chm14609
f1_keywords:
- vbaac10.chm14609
ms.prod: access
api_name:
- Access.Section.AlternateBackShade
ms.assetid: 0554bd30-1881-39c3-75ed-39d9164a7ae5
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.AlternateBackShade property (Access)

Gets or sets the shade applied to the theme color in the **AlternateBackColor** property of the section. Read/write **Single**.


## Syntax

_expression_.**AlternateBackShade**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

The **AlternateBackShade** property contains a numeric expression that can be used to darken the theme color in the **AlternateBackColor** property. The default value of the **AlternateBackShade** property is 100, which is neutral, and does not change the theme color. 

To darken the color, first determine the percentage by which to darken from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to darken the theme color shade by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]