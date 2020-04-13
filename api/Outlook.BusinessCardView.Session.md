---
title: BusinessCardView.Session property (Outlook)
keywords: vbaol11.chm2919
f1_keywords:
- vbaol11.chm2919
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Session
ms.assetid: 18e5fb02-1d57-3c47-74ed-0409d734b4cb
ms.date: 06/08/2017
localization_priority: Normal
---


# BusinessCardView.Session property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents a [BusinessCardView](Outlook.BusinessCardView.md) object.


## Remarks

The **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[BusinessCardView Object](Outlook.BusinessCardView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]