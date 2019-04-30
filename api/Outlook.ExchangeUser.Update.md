---
title: ExchangeUser.Update method (Outlook)
keywords: vbaol11.chm2076
f1_keywords:
- vbaol11.chm2076
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.Update
ms.assetid: a2672fbf-f32a-f120-227c-24ee5c361f35
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.Update method (Outlook)

Posts a change to the  **[ExchangeUser](Outlook.ExchangeUser.md)** object in the messaging system.


## Syntax

_expression_.**Update** (_MakePermanent_, _Refresh_)

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MakePermanent_|Optional| **Variant**|A value of  **True** indicates that the property cache is flushed and all changes are committed in the underlying address book. A value of **False** indicates that the property cache is flushed but not committed to persistent storage. The default value is **True**.|
| _Refresh_|Optional| **Variant**|A value of  **True** indicates that the property cache is reloaded from the values in the underlying address book. A value of **False** indicates that the property cache is not reloaded. The default value is **False**.|

## Remarks

New entries or changes to existing entries are not persisted in the collection until the  **Update** method has been called with its _MakePermanent_ parameter set to **True**. 

To flush the cache and then reload the values from the address book, call  **Update** with the _MakePermanent_ parameter set to **False** and the _Refresh_ parameter set to **True**. 


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]