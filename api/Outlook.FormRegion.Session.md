---
title: FormRegion.Session property (Outlook)
keywords: vbaol11.chm2387
f1_keywords:
- vbaol11.chm2387
ms.prod: outlook
api_name:
- Outlook.FormRegion.Session
ms.assetid: 13b9a148-c898-a3ef-8341-073767ce665e
ms.date: 06/08/2017
localization_priority: Normal
---


# FormRegion.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [FormRegion](Outlook.FormRegion.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[FormRegion Object](Outlook.FormRegion.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]