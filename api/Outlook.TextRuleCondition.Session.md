---
title: TextRuleCondition.Session Property (Outlook)
keywords: vbaol11.chm2474
f1_keywords:
- vbaol11.chm2474
ms.prod: outlook
api_name:
- Outlook.TextRuleCondition.Session
ms.assetid: 29422538-9045-66b5-44a1-b226870dc307
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRuleCondition.Session Property (Outlook)

Returns the  **[NameSpace](Outlook.NameSpace.md)** object for the current session. Read-only.


## Syntax

 _expression_. `Session`

_expression_ A variable that represents a [TextRuleCondition](./Outlook.TextRuleCondition.md) object.


## Remarks

The  **Session** property and the **[GetNamespace](Outlook.Application.GetNamespace.md)** method can be used interchangeably to obtain the **NameSpace** object for the current session. Both members serve the same purpose. For example, the following statements perform the same function:


```vb
Set objNamespace = Application.GetNamespace("MAPI") 
```


```vb
Set objSession = Application.Session
```


## See also


[TextRuleCondition Object](Outlook.TextRuleCondition.md)

