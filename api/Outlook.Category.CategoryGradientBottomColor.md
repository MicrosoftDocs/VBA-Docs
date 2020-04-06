---
title: Category.CategoryGradientBottomColor property (Outlook)
keywords: vbaol11.chm3268
f1_keywords:
- vbaol11.chm3268
ms.prod: outlook
api_name:
- Outlook.Category.CategoryGradientBottomColor
ms.assetid: 5f082300-2eb0-b297-dc54-9657da5ae319
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.CategoryGradientBottomColor property (Outlook)

Returns an **OLE_COLOR** value that represents the bottom gradient color of the color swatch displayed for a **[Category](Outlook.Category.md)** object. Read-only.


## Syntax

_expression_. `CategoryGradientBottomColor`

_expression_ A variable that represents a [Category](Outlook.Category.md) object.


## Remarks

Setting the  **[Color](Outlook.Category.Color.md)** property of the **Category** object to an **[OlCategoryColor](Outlook.OlCategoryColor.md)** constant changes the value of this property, as well as the value of the **[CategoryGradientTopColor](Outlook.Category.CategoryGradientTopColor.md)** and **[CategoryBorderColor](Outlook.Category.CategoryBorderColor.md)** properties.

This color should be used to display a gradient-shaded color swatch for the  **Category** object.


## See also


[Category Object](Outlook.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]