---
title: IconView.Session property (Outlook)
keywords: vbaol11.chm2559
f1_keywords:
- vbaol11.chm2559
ms.prod: outlook
api_name:
- Outlook.IconView.Session
ms.assetid: 456b7396-f69c-57bb-1e71-cfc26b9e5613
ms.date: 06/08/2017
localization_priority: Normal
---


# IconView.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [IconView](Outlook.IconView.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[IconView Object](Outlook.IconView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]