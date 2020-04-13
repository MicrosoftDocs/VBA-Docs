---
title: Inspector.Session property (Outlook)
keywords: vbaol11.chm2959
f1_keywords:
- vbaol11.chm2959
ms.prod: outlook
api_name:
- Outlook.Inspector.Session
ms.assetid: e3e36957-1df2-af40-83e7-c5825ceb9c4d
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]