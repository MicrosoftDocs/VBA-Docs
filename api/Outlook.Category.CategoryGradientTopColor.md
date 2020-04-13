---
title: Category.CategoryGradientTopColor property (Outlook)
keywords: vbaol11.chm3267
f1_keywords:
- vbaol11.chm3267
ms.prod: outlook
api_name:
- Outlook.Category.CategoryGradientTopColor
ms.assetid: deb7a986-8afd-465c-ed8e-3cf669f96a35
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.CategoryGradientTopColor property (Outlook)

Returns an **OLE_COLOR** value that represents the top gradient color of the color swatch displayed for a **[Category](Outlook.Category.md)** object. Read-only.


## Syntax

_expression_. `CategoryGradientTopColor`

_expression_ A variable that represents a [Category](Outlook.Category.md) object.


## Remarks

Setting the  **[Color](Outlook.Category.Color.md)** property of the **Category** object to an **[OlCategoryColor](Outlook.OlCategoryColor.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientBottomColor](Outlook.Category.CategoryGradientBottomColor.md)** and **[CategoryBorderColor](Outlook.Category.CategoryBorderColor.md)** properties.

This color should be used to display a gradient-shaded color swatch for the  **Category** object.


## See also


[Category Object](Outlook.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]