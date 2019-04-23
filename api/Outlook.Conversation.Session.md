---
title: Conversation.Session property (Outlook)
keywords: vbaol11.chm3386
f1_keywords:
- vbaol11.chm3386
ms.prod: outlook
api_name:
- Outlook.Conversation.Session
ms.assetid: 6f41faaa-e16a-d171-ed72-d2fef64a77f1
ms.date: 06/08/2017
localization_priority: Normal
---


# Conversation.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a '[Conversation](Outlook.Conversation.md)' object.


## Remarks

This property returns  **Null** (**Nothing** in Visual Basic) if there is no logged-on session.

You can use the  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:




```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```




```vb
Set objSession = Application.Session
```


## See also


[Conversation Object](Outlook.Conversation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]