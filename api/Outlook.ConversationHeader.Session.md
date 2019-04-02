---
title: ConversationHeader.Session property (Outlook)
keywords: vbaol11.chm3548
f1_keywords:
- vbaol11.chm3548
ms.prod: outlook
api_name:
- Outlook.ConversationHeader.Session
ms.assetid: 1262a068-ad5f-492d-2a96-edc365956fe6
ms.date: 06/08/2017
localization_priority: Normal
---


# ConversationHeader.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a '[ConversationHeader](Outlook.ConversationHeader.md)' object.


## Remarks

Returns  **Null** (**Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method interchangeably to obtain the **NameSpace** object for the current session.


## See also


[ConversationHeader Object](Outlook.ConversationHeader.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]