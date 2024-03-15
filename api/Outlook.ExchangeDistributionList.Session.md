---
title: ExchangeDistributionList.Session property (Outlook)
keywords: vbaol11.chm2110
f1_keywords:
- vbaol11.chm2110
api_name:
- Outlook.ExchangeDistributionList.Session
ms.assetid: 9488e161-d297-d999-538d-a8b295380701
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# ExchangeDistributionList.Session property (Outlook)

Returns the **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

_expression_.**Session**

_expression_ A variable that represents an [ExchangeDistributionList](Outlook.ExchangeDistributionList.md) object.


## Remarks

The **Session** property and the **[Application.GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[ExchangeDistributionList Object](Outlook.ExchangeDistributionList.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]