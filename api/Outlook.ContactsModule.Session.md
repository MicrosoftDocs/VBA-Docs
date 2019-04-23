---
title: ContactsModule.Session property (Outlook)
keywords: vbaol11.chm2834
f1_keywords:
- vbaol11.chm2834
ms.prod: outlook
api_name:
- Outlook.ContactsModule.Session
ms.assetid: 4ab5d6e1-fcff-9aa4-0779-a365e01d0a1c
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactsModule.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [ContactsModule](Outlook.ContactsModule.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ContactsModule Object](Outlook.ContactsModule.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]