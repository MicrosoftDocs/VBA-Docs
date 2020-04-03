---
title: PropertyAccessor.UTCToLocalTime method (Outlook)
keywords: vbaol11.chm1974
f1_keywords:
- vbaol11.chm1974
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.UTCToLocalTime
ms.assetid: a56311ac-60ac-4f51-5255-d6840bf6004d
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.UTCToLocalTime method (Outlook)

Converts the date-time value that is specified by  _Value_ and expressed in Coordinated Universal Time (UTC) to a value in local time.


## Syntax

_expression_. `UTCToLocalTime`( `_Value_` )

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **Date**|The date-time value to be converted from UTC to local time.|

## Return value

A **Date** value that represents _Value_ after being converted from UTC to local time.


## Remarks

For more information on type conversion when using the  **[PropertyAccessor](Outlook.PropertyAccessor.md)** object, see [Best Practices for Getting and Setting Properties](../outlook/How-to/Navigation/best-practices-for-getting-and-setting-properties.md).


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]