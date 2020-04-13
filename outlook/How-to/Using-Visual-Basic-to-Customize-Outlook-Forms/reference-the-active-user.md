---
title: Reference the Active User
keywords: olfm10.chm3077415
f1_keywords:
- olfm10.chm3077415
ms.prod: outlook
ms.assetid: dc8e5e24-51a0-1b16-649e-6b24e0fa9b56
ms.date: 06/08/2019
localization_priority: Normal
---


# Reference the Active User

Use **[Application.GetNamespace](../../../api/Outlook.Application.GetNamespace.md)** to return the Outlook **[NameSpace](../../../api/Outlook.NameSpace.md)** object from the **[Application](../../../api/Outlook.Application.md)** object, and then use the **[NameSpace.CurrentUser](../../../api/Outlook.NameSpace.CurrentUser.md)** property to return a **[Recipient](../../../api/Outlook.Recipient.md)** object representing the active user, as shown in the following example.


```vb
Set myUser = Application.GetNameSpace("MAPI").CurrentUser
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]