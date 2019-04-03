---
title: ExchangeUser.DisplayType property (Outlook)
keywords: vbaol11.chm2066
f1_keywords:
- vbaol11.chm2066
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.DisplayType
ms.assetid: 3060a00b-9a99-7833-1a8a-5c18123d7d98
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.DisplayType property (Outlook)

Returns  **olUser** which is a constant from the **[OlDisplayType](Outlook.OlDisplayType.md)** enumeration representing the nature of the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read-only.


## Syntax

_expression_.**DisplayType**

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property corresponds to the MAPI property,  **PidTagDisplayType**.

 The **ExchangeUser** object is derived from the **[AddressEntry](Outlook.AddressEntry.md)** object. It inherits the **DisplayType** property from the **AddressEntry** object. In the case of **ExchangeUser**, **DisplayType** should always return **olUser**.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]