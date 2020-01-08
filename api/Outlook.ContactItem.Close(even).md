---
title: ContactItem.Close event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.Close
ms.assetid: beeeb53c-94fe-ae1b-7870-87bd37b3debf
ms.date: 12/17/2019
localization_priority: Normal
---


# ContactItem.Close event (Outlook)

Occurs when the inspector associated with an item (which is an instance of the parent object) is being closed.


## Syntax

_expression_.**Close** (_Cancel_)

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the close operation is not completed and the inspector is left open.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the close operation isn't completed and the inspector is left open.

If you use the  **[Close](Outlook.ContactItem.Close(method).md)** method to fire this event, it can only be canceled if the **Close** method uses the **olPromptForSave** argument.

>[!NOTE]
> In C#, there is no ContactItem.Close event, only [Close](https://docs.microsoft.com/office/vba/api/outlook.contactitem.close(method)).
>
> Try the following code example instead:
>`ItemEvents_10_Event appointmentItemEvent = (ItemEvents_10_Event)outlook_contact;
appointmentItemEvent.Close += AppointmentItemEvent_Close;`


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
