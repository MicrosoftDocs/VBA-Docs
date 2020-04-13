---
title: Category.CategoryBorderColor property (Outlook)
keywords: vbaol11.chm3266
f1_keywords:
- vbaol11.chm3266
ms.prod: outlook
api_name:
- Outlook.Category.CategoryBorderColor
ms.assetid: 95251459-f216-7cc8-55ef-c939090cf3bf
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.CategoryBorderColor property (Outlook)

Returns an **OLE_COLOR** value that represents the border color of the color swatch displayed for a **[Category](Outlook.Category.md)** object. Read-only.


## Syntax

_expression_. `CategoryBorderColor`

_expression_ A variable that represents a [Category](Outlook.Category.md) object.


## Remarks

Setting the  **[Color](Outlook.Category.Color.md)** property of the **Category** object to an **[OlCategoryColor](Outlook.OlCategoryColor.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientBottomColor](Outlook.Category.CategoryGradientBottomColor.md)** and **[CategoryGradientTopColor](Outlook.Category.CategoryGradientTopColor.md)** properties.

This color should be used to display the border of a gradient-shaded color swatch for the  **Category** object.


## See also


[Category Object](Outlook.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]