---
title: ViewFields.Insert method (Outlook)
keywords: vbaol11.chm2553
f1_keywords:
- vbaol11.chm2553
ms.prod: outlook
api_name:
- Outlook.ViewFields.Insert
ms.assetid: a975a030-76c9-e877-8df7-601094998fd1
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFields.Insert method (Outlook)

Creates a new **[ViewField](Outlook.ViewField.md)** object and inserts it at the specified index within the **[ViewFields](Outlook.ViewFields.md)** collection.


## Syntax

_expression_.**Insert** (_PropertyName_, _Index_)

_expression_ A variable that represents a [ViewFields](Outlook.ViewFields.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required| **String**|The name of the property to which the new object is associated.|
| _Index_|Required| **Variant**|Either a one-based index number at which to insert the new object, or a value used to match the  **[ViewXMLSchemaName](Outlook.ViewField.ViewXMLSchemaName.md)** property value of an object in the collection where the new object is to be inserted.|

## Return value

A  **ViewField** object that represents the new view field.


## See also


[ViewFields Object](Outlook.ViewFields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]