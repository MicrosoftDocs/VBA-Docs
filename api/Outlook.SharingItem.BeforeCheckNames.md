---
title: SharingItem.BeforeCheckNames event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeCheckNames
ms.assetid: f6e787a5-dad1-6d4e-119f-1d0f07133a83
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.BeforeCheckNames event (Outlook)

Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).


## Syntax

_expression_. `BeforeCheckNames`( `_Cancel_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the name resolution process is not completed.|

## Remarks

You use the  **BeforeCheckNames** event in VBScript, but the event does not fire when an email name is resolved on the form.

The event does not fire under the following circumstances:


- You customized a Journal Entry form and then resolved a contact in the  **Contacts** field.
    
- You customized a Contact form and then resolved a contact in the  **Contacts** field.
    
- You customized any type of form and Outlook automatically resolved the name in the background.
    
- You programmatically created and resolved a recipient.
    



## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]