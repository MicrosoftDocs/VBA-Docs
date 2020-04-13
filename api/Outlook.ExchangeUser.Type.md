---
title: ExchangeUser.Type property (Outlook)
keywords: vbaol11.chm2072
f1_keywords:
- vbaol11.chm2072
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Type
ms.assetid: de3652a8-023c-5d2c-9ced-88f768c22a87
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Type property (Outlook)

Returns a **String** representing the type of entry for the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read/write.


## Syntax

_expression_.**Type**

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

The **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **Type** property from the **AddressEntry** object. In the case of **ExchangeUser**, the type is always the string "EX".


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]