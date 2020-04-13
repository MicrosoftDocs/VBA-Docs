---
title: OutlookBarStorage.Session property (Outlook)
keywords: vbaol11.chm370
f1_keywords:
- vbaol11.chm370
ms.prod: outlook
api_name:
- Outlook.OutlookBarStorage.Session
ms.assetid: f3ba6302-aca2-f8ba-3a82-ae35f6b5b609
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarStorage.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OutlookBarStorage](Outlook.OutlookBarStorage.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OutlookBarStorage Object](Outlook.OutlookBarStorage.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]