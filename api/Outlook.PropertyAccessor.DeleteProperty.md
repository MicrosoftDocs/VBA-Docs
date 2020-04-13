---
title: PropertyAccessor.DeleteProperty method (Outlook)
keywords: vbaol11.chm1978
f1_keywords:
- vbaol11.chm1978
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.DeleteProperty
ms.assetid: 9acb52b5-13a7-7363-7e17-83804037f33b
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.DeleteProperty method (Outlook)

Deletes the property specified by  _SchemaName_.


## Syntax

_expression_. `DeleteProperty`( `_SchemaName_` )

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SchemaName_|Required| **String**|The name of the property that is to be deleted for the parent object of the  **[PropertyAccessor](Outlook.PropertyAccessor.md)** object. The property is referenced by namespace. For more information, see [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md).|

## Remarks

The caller must have the permission to delete properties. The **DeleteProperty** method deletes only custom properties; it does not delete any Outlook built-in property or any MAPI property. It does not delete custom properties of the **[DocumentItem](Outlook.DocumentItem.md)** object.


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]