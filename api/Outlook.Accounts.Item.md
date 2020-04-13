---
title: Accounts.Item method (Outlook)
keywords: vbaol11.chm750
f1_keywords:
- vbaol11.chm750
ms.prod: outlook
api_name:
- Outlook.Accounts.Item
ms.assetid: 8ef9c358-6d8b-1cbb-40ed-6d3462ae335e
ms.date: 06/08/2017
localization_priority: Normal
---


# Accounts.Item method (Outlook)

Returns an **[Account](Outlook.Account.md)** object specified by _Index_.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an [Accounts](Outlook.Accounts.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|A one-based  **Long** that indexes into the **[Accounts](Outlook.Accounts.md)** collection, or a **String** that specifies the **[DisplayName](Outlook.Account.DisplayName.md)** of an **Account**.|

## Return value

An **Account** object that matches the account specified by _Index_.


## See also


[Accounts Object](Outlook.Accounts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]