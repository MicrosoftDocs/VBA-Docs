---
title: ExchangeUser.AssistantName property (Outlook)
keywords: vbaol11.chm2086
f1_keywords:
- vbaol11.chm2086
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.AssistantName
ms.assetid: cca35d99-7031-ba46-4171-8c89b9ea2e2b
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.AssistantName property (Outlook)

Returns a  **String** representing the name of the assistant for the **[ExchangeUser](Outlook.ExchangeUser.md)**. Read/write.


## Syntax

_expression_. `AssistantName`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Remarks

This property corresponds to MAPI property  **PidTagAssistant**.

Returns an empty string if this property has not been implemented or does not exist for the  **ExchangeUser** object.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]