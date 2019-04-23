---
title: PropertyAccessor.LocalTimeToUTC method (Outlook)
keywords: vbaol11.chm1975
f1_keywords:
- vbaol11.chm1975
ms.prod: outlook
api_name:
- Outlook.PropertyAccessor.LocalTimeToUTC
ms.assetid: c19f60b2-441f-77b3-eb83-9cfd899e3a52
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyAccessor.LocalTimeToUTC method (Outlook)

Converts a date-time value specified by  _Value_ from the local time format to Coordinated Universal Time (UTC) format.


## Syntax

_expression_. `LocalTimeToUTC`( `_Value_` )

_expression_ A variable that represents a [PropertyAccessor](Outlook.PropertyAccessor.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Value_|Required| **Date**|The date-time value to be converted from local time to UTC.|

## Return value

A  **Date** value that represents _Value_ after being converted from local time to UTC.


## Remarks

For more information on type conversion when using the  **PropertyAccessor** object, see [Best Practices for Getting and Setting Properties](../outlook/How-to/Navigation/best-practices-for-getting-and-setting-properties.md).


## See also


[PropertyAccessor Object](Outlook.PropertyAccessor.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]