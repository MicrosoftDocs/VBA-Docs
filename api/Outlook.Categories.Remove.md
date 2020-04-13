---
title: Categories.Remove method (Outlook)
keywords: vbaol11.chm2438
f1_keywords:
- vbaol11.chm2438
ms.prod: outlook
api_name:
- Outlook.Categories.Remove
ms.assetid: 8c16b02e-0297-9f36-7cb7-20e6ab0c286b
ms.date: 06/08/2017
localization_priority: Normal
---


# Categories.Remove method (Outlook)

Removes an object from the collection.


## Syntax

_expression_.**Remove** (_Index_)

_expression_ A variable that represents a [Categories](Outlook.Categories.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a **Long** value representing the index number of the object, or a **String** value representing either the **[Name](Outlook.Category.Name.md)** or **[CategoryID](Outlook.Category.CategoryID.md)** property value of an object in the collection.|

## Remarks

If the name of a category is specified in  _Index_, this method removes the first  **Category** object that matches the specified value.


## See also


[Categories Object](Outlook.Categories.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]