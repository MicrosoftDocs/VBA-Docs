---
title: Stores.BeforeStoreRemove event (Outlook)
keywords: vbaol11.chm2754
f1_keywords:
- vbaol11.chm2754
ms.prod: outlook
api_name:
- Outlook.Stores.BeforeStoreRemove
ms.assetid: b21d4854-3da5-5c01-cbc1-098bb505466e
ms.date: 06/08/2017
localization_priority: Normal
---


# Stores.BeforeStoreRemove event (Outlook)

Occurs when a **[Store](Outlook.Store.md)** is about to be removed from the current session either programmatically or through user action.


## Syntax

_expression_. `BeforeStoreRemove`( `_Store_` , `_Cancel_` )

_expression_ A variable that represents a [Stores](Outlook.Stores.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **Store**|The **Store** object to be removed from the current session.|
| _Cancel_|Required| **Boolean**| **True** to cancel the removal of the specified store, **False** otherwise.|

## Remarks

Outlook must be running in order for this event to fire. This event will fire when any of the following occurs:


- A store is removed by the user clicking the  **Close** command on the Shortcut menu.
    
- A store is removed programmatically by calling  **[Namespace.RemoveStore](Outlook.NameSpace.RemoveStore.md)**.
    


This event will not fire when any of the following occurs:


- When Outlook shuts down and closes a primary or delegate store.
    
- If a store is removed through the  **Mail** applet in the Microsoft Windows Control Panel and Outlook is not running.
    
- A delegate store is removed on the  **Advanced** tab of the **Microsoft Exchange Server** dialog box.
    
- A store is removed through the  **Data Files** tab of the **Account Manager** dialog box when Outlook is not running.
    
- An IMAP Store is removed from the profile.
    


You can use this event to determine that a store has been removed, and take appropriate actions if the store is required for your application (such as remounting the store). Otherwise you would have to resort to polling the  **[Stores](Outlook.Stores.md)** collection.


## See also


[Stores Object](Outlook.Stores.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]