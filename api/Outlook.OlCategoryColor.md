---
title: OlCategoryColor enumeration (Outlook)
keywords: vbaol11.chm3119
f1_keywords:
- vbaol11.chm3119
ms.prod: outlook
api_name:
- Outlook.OlCategoryColor
ms.assetid: 048bbc6b-c49f-68a3-ac59-b61204e5ef78
ms.date: 06/08/2017
localization_priority: Normal
---


# OlCategoryColor enumeration (Outlook)

Indicates the color that is specified for a category or a font in a view.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olCategoryColorBlack**|15|Black|
| **olCategoryColorBlue**|8|Blue|
| **olCategoryColorDarkBlue**|23|Dark Blue|
| **olCategoryColorDarkGray**|14|Dark Gray|
| **olCategoryColorDarkGreen**|20|Dark Green|
| **olCategoryColorDarkMaroon**|25|Dark Maroon|
| **olCategoryColorDarkOlive**|22|Dark Olive|
| **olCategoryColorDarkOrange**|17|Dark Orange|
| **olCategoryColorDarkPeach**|18|Dark Peach|
| **olCategoryColorDarkPurple**|24|Dark Purple|
| **olCategoryColorDarkRed**|16|Dark Red|
| **olCategoryColorDarkSteel**|12|Dark Steel|
| **olCategoryColorDarkTeal**|21|Dark Teal|
| **olCategoryColorDarkYellow**|19|Dark Yellow|
| **olCategoryColorGray**|13|Gray|
| **olCategoryColorGreen**|5|Green|
| **olCategoryColorMaroon**|10|Maroon|
| **olCategoryColorNone**|0|No color assigned.|
| **olCategoryColorOlive**|7|Olive|
| **olCategoryColorOrange**|2|Orange|
| **olCategoryColorPeach**|3|Peach|
| **olCategoryColorPurple**|9|Purple|
| **olCategoryColorRed**|1|Red|
| **olCategoryColorSteel**|11|Steel|
| **olCategoryColorTeal**|6|Teal|
| **olCategoryColorYellow**|4|Yellow|

## Remarks

Used by the [Color](Outlook.Category.Color.md) property of the [Category object (Outlook)](Outlook.Category.md), and the [ExtendedColor](Outlook.ViewFont.ExtendedColor.md) property of the [ViewFont object (Outlook)](Outlook.ViewFont.md).

The color constants provided here are approximations of the actual colors used by the  **Category** object. Use the [CategoryBorderColor](Outlook.Category.CategoryBorderColor.md), [CategoryGradientBottomColor](Outlook.Category.CategoryGradientBottomColor.md), and [CategoryGradientTopColor](Outlook.Category.CategoryGradientTopColor.md) properties to retrieve the **OLE_COLOR** color values that are used to represent the **Category** object, after setting the **Color** property to the appropriate constant.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]