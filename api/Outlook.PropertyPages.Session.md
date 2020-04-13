---
title: PropertyPages.Session property (Outlook)
keywords: vbaol11.chm163
f1_keywords:
- vbaol11.chm163
ms.prod: outlook
api_name:
- Outlook.PropertyPages.Session
ms.assetid: 0a6c6235-b27b-72d4-bd17-c94627b91d41
ms.date: 06/08/2017
localization_priority: Normal
---


# PropertyPages.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [PropertyPages](Outlook.PropertyPages.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[PropertyPages Object](Outlook.PropertyPages.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]