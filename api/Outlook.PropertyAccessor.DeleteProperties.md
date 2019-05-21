---
title: PropertyAccessor.DeleteProperties method (Outlook)
keywords: vbaol11.chm1979
f1_keywords:
- vbaol11.chm1979
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.DeleteProperties
ms.assetid: e9c11799-cb75-fd8c-0c98-aca46796bb46
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.DeleteProperties method (Outlook)

Deletes the properties specified in the array  _SchemaNames_.


## Syntax

_expression_. `DeleteProperties`( `_SchemaNames_` )

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SchemaNames_|Required| **Variant**|An array that contains the names of the properties that are to be deleted for the parent object of the  **[PropertyAccessor](Outlook.PropertyAccessor.md)** object. These properties are referenced by namespace. For more information, see [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md).|

## Return value

A Variant that is  **Null** (**Nothing** in VBA) if the operation is successful, or is an array of **[Err](../language/reference/User-Interface-Help/err-object.md)** objects if an error occurs. If the return value is an array, the size of this array is the same as that of the _SchemaNames_ array. An **Err** value in the array is mapped to the error result of deleting the corresponding property in the _SchemaNames_ parameter.


## Remarks

The caller must have the permission to delete properties. The  **DeleteProperties** method only deletes custom properties that exist. It does not delete any Outlook built-in property or any MAPI property. It does not delete custom properties of the **[DocumentItem](Outlook.DocumentItem.md)** object.


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]