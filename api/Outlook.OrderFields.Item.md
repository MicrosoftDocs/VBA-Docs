---
title: OrderFields.Item method (Outlook)
keywords: vbaol11.chm2677
f1_keywords:
- vbaol11.chm2677
ms.prod: outlook
api_name:
- Outlook.OrderFields.Item
ms.assetid: 0738f59e-8eda-18af-1aee-13d566c248db
ms.date: 06/08/2017
localization_priority: Normal
---


# OrderFields.Item method (Outlook)

Returns an  **[OrderField](Outlook.OrderField.md)** object from the collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an [OrderFields](Outlook.OrderFields.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The value can be a one-based integer that indexes an  **OrderField** object in the **[OrderFields](Outlook.OrderFields.md)** collection, a string that matches the **[ViewXMLSchemaName](Outlook.OrderField.ViewXMLSchemaName.md)** property value of an **OrderField** object in the collection, or a field name as displayed in the Field Chooser.|

## Return value

An  **OrderField** object that represents the specified object.


## See also


[OrderFields Object](Outlook.OrderFields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]