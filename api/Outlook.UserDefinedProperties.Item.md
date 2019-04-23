---
title: UserDefinedProperties.Item method (Outlook)
keywords: vbaol11.chm587
f1_keywords:
- vbaol11.chm587
ms.prod: outlook
api_name:
- Outlook.UserDefinedProperties.Item
ms.assetid: 45f5ec00-00c6-2e90-68bc-6bcab79cada6
ms.date: 06/08/2017
localization_priority: Normal
---


# UserDefinedProperties.Item method (Outlook)

Returns an object from the collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [UserDefinedProperties](Outlook.UserDefinedProperties.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either a  **Long** value that represents the 1-based index number of an object in the collection, or a **String** value that represents the **[Name](Outlook.UserDefinedProperty.Name.md)** property value of an object in the collection.|

## Return value

A  **[UserDefinedProperty](Outlook.UserDefinedProperty.md)** object that represents the specified object.


## See also


[UserDefinedProperties Object](Outlook.UserDefinedProperties.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]