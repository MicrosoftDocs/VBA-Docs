---
title: OutlookBarPane.Session property (Outlook)
keywords: vbaol11.chm361
f1_keywords:
- vbaol11.chm361
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.Session
ms.assetid: 8aa3d36b-2044-85a7-2b79-86c6b161c4df
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlookBarPane.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [OutlookBarPane](Outlook.OutlookBarPane.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements do the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[OutlookBarPane Object](Outlook.OutlookBarPane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]