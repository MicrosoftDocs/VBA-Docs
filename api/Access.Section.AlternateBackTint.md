---
title: Section.AlternateBackTint property (Access)
keywords: vbaac10.chm14608
f1_keywords:
- vbaac10.chm14608
ms.prod: access
api_name:
- Access.Section.AlternateBackTint
ms.assetid: 7758713d-cfba-ac57-91c7-fcdab26ae44a
ms.date: 03/23/2019
localization_priority: Normal
---


# Section.AlternateBackTint property (Access)

Gets or sets the tint applied to the theme color in the **AlternateBackColor** property of the section. Read/write **Single**.


## Syntax

_expression_.**AlternateBackTint**

_expression_ A variable that represents a **[Section](Access.Section.md)** object.


## Remarks

The **AlternateBackTint** property contains a numeric expression that can be used to lighten the theme color in the **AlternateBackColor** property. The default value of the **AlternateBackTint** property is 100, which is neutral, and does not change the theme color. 

To lighten the color, first determine the percentage by which to lighten from 1 to 100, and then subtract that value as a whole number from 100 and use the remainder. For example, to lighten the theme color tint by 75%, subtract 75 from 100 and use the remainder, which is 25.

This property is not surfaced in the property sheet.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]