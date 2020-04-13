---
title: Stores.Item method (Outlook)
keywords: vbaol11.chm819
f1_keywords:
- vbaol11.chm819
ms.prod: outlook
api_name:
- Outlook.Stores.Item
ms.assetid: b516241a-7baf-b04b-027d-25de80058fbe
ms.date: 06/08/2017
localization_priority: Normal
---


# Stores.Item method (Outlook)

Returns a **[Store](Outlook.Store.md)** object that is specified by _Index_. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a [Stores](Outlook.Stores.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|Either an **Integer** that specifies a one-based index into the **Stores** collection, or a **String** value that specifies the **[DisplayName](Outlook.Store.DisplayName.md)** of a **Store** in the **Stores** collection.|

## Return value

A **Store** object in the parent **[Stores](Outlook.Stores.md)** collection, as specified by _Index_.


## Remarks

The **Store.DisplayName** property is the default property of a **Store**.

If  _Index_ is a string and no item can be found by that name, an error will be returned.


## See also


[Stores Object](Outlook.Stores.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]