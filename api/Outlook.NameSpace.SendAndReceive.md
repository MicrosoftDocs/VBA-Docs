---
title: NameSpace.SendAndReceive method (Outlook)
keywords: vbaol11.chm782
f1_keywords:
- vbaol11.chm782
ms.prod: outlook
api_name:
- Outlook.NameSpace.SendAndReceive
ms.assetid: 196b15b0-6766-ca2a-8dbe-991fc93b8307
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.SendAndReceive method (Outlook)

Initiates immediate delivery of all undelivered messages submitted in the current session, and immediate receipt of mail for all accounts in the current profile. 


## Syntax

_expression_. `SendAndReceive`( `_showProgressDialog_` )

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _showProgressDialog_|Required| **Boolean**|Indicates whether the  **Outlook Send/Receive Progress** dialog box should be displayed, regardless of user settings.|

## Remarks

Calling the  **SendAndReceive** method is asynchronous.

 **SendAndReceive** provides the programmatic equivalent to the **Send/Receive All** command that is available when you click **Tools** and then **Send/Receive**.

If you do not need to synchronize all objects, you can use the  **[SyncObjects](Outlook.SyncObjects.md)** collection object to select specific objects. For more information, see **[NameSpace.SyncObjects](Outlook.NameSpace.SyncObjects.md)**.

All accounts defined in the current profile are used in  **Send/Receive All**. If an online connection is required to perform the  **Send/Receive All**, the connection is made according to user preferences.


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]