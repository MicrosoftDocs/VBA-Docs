---
title: SharingItem.Send Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Send
ms.assetid: 15db902f-d61d-cfcd-0498-a2cec5f984bb
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Send Event (Outlook)

Occurs when the user selects the  **Send** action for an item, or when the **[Send](Outlook.SharingItem.Send(method).md)** method is called for the item, which is an instance of the parent object.


## Syntax

_expression_. `Send`( `_Cancel_` )

 _expression_ An expression that returns a [SharingItem](./Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the send operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the item is not sent.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]