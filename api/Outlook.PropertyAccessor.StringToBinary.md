---
title: PropertyAccessor.StringToBinary method (Outlook)
keywords: vbaol11.chm1976
f1_keywords:
- vbaol11.chm1976
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.StringToBinary
ms.assetid: 1ea95601-a21f-47d2-7a3c-166c4984fc25
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.StringToBinary method (Outlook)

Converts a string specified by  _Value_ to an array of bytes.


## Syntax

_expression_. `StringToBinary`( `_Value_` )

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **String**|A hexadecimal string value that is to be converted to an array of bytes.|

## Return value

A  **Variant** value that represents an array of bytes returned from the conversion.


## Remarks

For more information on type conversion when using the  **PropertyAccessor** object, see [Best Practices for Getting and Setting Properties](../outlook/How-to/Navigation/best-practices-for-getting-and-setting-properties.md).


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]