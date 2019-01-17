---
title: TaskRequestAcceptItem.BeforeCheckNames Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestAcceptItem.BeforeCheckNames
ms.assetid: 152328cc-8ef9-5222-8644-7ea27960bed4
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestAcceptItem.BeforeCheckNames Event (Outlook)

Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).


## Syntax

_expression_. `BeforeCheckNames`( `_Cancel_` )

_expression_ A variable that represents a [TaskRequestAcceptItem](./Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the name resolution process is not completed.|

## Remarks

You use the  **BeforeCheckNames** event in VBScript, but the event does not fire when an email name is resolved on the form.

The event does not fire under the following circumstances:


- You customized a Journal Entry form and then resolved a contact in the  **Contacts** field.
    
- You customized a Contact form and then resolved a contact in the  **Contacts** field.
    
- You customized any type of form and Outlook automatically resolved the name in the background.
    
- You programmatically created and resolved a recipient.
    



## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

