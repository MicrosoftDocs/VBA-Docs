---
title: Actions.Session property (Outlook)
keywords: vbaol11.chm147
f1_keywords:
- vbaol11.chm147
ms.prod: outlook
api_name:
- Outlook.Actions.Session
ms.assetid: 21792c3f-9669-2f68-7a47-bac172d16620
ms.date: 06/08/2017
localization_priority: Normal
---


# Actions.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [Actions](Outlook.Actions.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[Actions Object](Outlook.Actions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]