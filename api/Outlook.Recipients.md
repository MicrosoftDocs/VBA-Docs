---
title: Recipients object (Outlook)
keywords: vbaol11.chm225
f1_keywords:
- vbaol11.chm225
ms.prod: outlook
api_name:
- Outlook.Recipients
ms.assetid: 774f56b7-4de8-9584-60cd-4fbf361f4c85
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipients object (Outlook)

Contains a collection of  **[Recipient](Outlook.Recipient.md)** objects for an Outlook item.


## Remarks

Use the  **Recipients** property to return the **Recipients** object of an **[AppointmentItem](Outlook.AppointmentItem.md)**, **[JournalItem](Outlook.JournalItem.md)**, **[MailItem](Outlook.MailItem.md)**, **[MeetingItem](Outlook.MeetingItem.md)**, or **[TaskItem](Outlook.TaskItem.md)** object.

Use the  **[Add](Outlook.Recipients.Add.md)** method to create a new **Recipient** object and add it to the **Recipients** object. The **[Type](Outlook.Recipient.Type.md)** property of a new **Recipient** object is set to the default for the associated **AppointmentItem**, **JournalItem**, **MailItem**, or **TaskItem** object and must be reset to indicate another recipient type.

Use  **Recipients** (_index_), where _index_ is the name or index number, to return a single **Recipient** object. The name can be a string representing the display name, the alias, or the full SMTP email address of the recipient.


## Example

The following example creates a new **MailItem** object and adds Jon Grande as the recipient using the default type ("To").


```vb
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande")
```

The following example creates the same  **MailItem** object as the preceding example, and then changes the type of the **Recipient** object from the default ("To") to CC.




```vb
Set myItem = Application.CreateItem(olMailItem) 
 
Set myRecipient = myItem.Recipients.Add ("Jon Grande") 
 
myRecipient.Type = olCC
```


## Methods



|Name|
|:-----|
|[Add](Outlook.Recipients.Add.md)|
|[Item](Outlook.Recipients.Item.md)|
|[Remove](Outlook.Recipients.Remove.md)|
|[ResolveAll](Outlook.Recipients.ResolveAll.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.Recipients.Application.md)|
|[Class](Outlook.Recipients.Class.md)|
|[Count](Outlook.Recipients.Count.md)|
|[Parent](Outlook.Recipients.Parent.md)|
|[Session](Outlook.Recipients.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)
[Recipients Object Members](overview/Outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
