---
title: ExchangeUser.CompanyName property (Outlook)
keywords: vbaol11.chm2090
f1_keywords:
- vbaol11.chm2090
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.CompanyName
ms.assetid: d7a630ec-0fbf-78ea-5f2a-51be6d001c23
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.CompanyName property (Outlook)

Returns a  **String** representing the name of the company for the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read/write.


## Syntax

_expression_. `CompanyName`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property corresponds to the MAPI property,  **PidTagCompanyName**.

 Returns an empty string if this property has not been implemented or does not exist for the **ExchangeUser** object.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]