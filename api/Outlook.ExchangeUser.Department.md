---
title: ExchangeUser.Department property (Outlook)
keywords: vbaol11.chm2091
f1_keywords:
- vbaol11.chm2091
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Department
ms.assetid: 3b2512ff-d741-53b2-6f1d-a0f74ffbbce1
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Department property (Outlook)

Returns a **String** representing the department for the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read/write.


## Syntax

_expression_. `Department`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property corresponds to the MAPI property,  **PidTagDepartmentName**.

 Returns an empty string if this property has not been implemented or does not exist for the **ExchangeUser** object.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]