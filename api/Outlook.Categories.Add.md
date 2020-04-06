---
title: Categories.Add method (Outlook)
keywords: vbaol11.chm2437
f1_keywords:
- vbaol11.chm2437
ms.prod: outlook
api_name:
- Outlook.Categories.Add
ms.assetid: f776c2a2-1b32-f4eb-de5e-6e245a60cac2
ms.date: 06/08/2017
localization_priority: Normal
---


# Categories.Add method (Outlook)

Creates a new  **[Category](Outlook.Category.md)** object and appends it to the **[Categories](Outlook.Categories.md)** collection.


## Syntax

_expression_.**Add** (_Name_, _Color_, _ShortcutKey_)

_expression_ A variable that represents a [Categories](Outlook.Categories.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new category.|
| _Color_|Optional| **[OlCategoryColor](Outlook.OlCategoryColor.md)**|The color for the new category. If no value is specified, the new category is set to the first color (as specified in the order of the  **OlCategoryColor** enumeration) that is the least used, That is, if there are unused colors, the new category is set to the first unused color in the **OlCategoryColor** enumeration. If all colors in the **OlCategoryColor** enumeration have been used, then the new category is set to the first color that is used least in the **OlCategoryColor** enumeration.|
| _ShortcutKey_|Optional| **[OlCategoryShortcutKey](Outlook.OlCategoryShortcutKey.md)**|The shortcut key for the new category. If no value is specified, the default value is  **OlCategoryShortcutKeyNone**.|

## Return value

A  **Category** object that represents the new category.


## See also


[Categories Object](Outlook.Categories.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]