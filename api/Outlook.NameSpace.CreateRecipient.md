---
title: NameSpace.CreateRecipient method (Outlook)
keywords: vbaol11.chm760
f1_keywords:
- vbaol11.chm760
ms.prod: outlook
api_name:
- Outlook.NameSpace.CreateRecipient
ms.assetid: 7134c0d7-5f60-c63c-2dde-492d52b78fbe
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.CreateRecipient method (Outlook)

Creates a  **[Recipient](Outlook.Recipient.md)** object.


## Syntax

_expression_. `CreateRecipient`( `_RecipientName_` )

_expression_ A variable that represents a [NameSpace](Outlook.NameSpace.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RecipientName_|Required| **String**|The name of the recipient; it can be a string representing the display name, the alias, or the full SMTP email address of the recipient.|

## Return value

A  **Recipient** object that represents the new recipient.


## Remarks

 This method is most commonly used to create a **Recipient** object for use with the **[GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md)** method, for example, to open a delegator's folder. It can also be used to verify a given name against an address book.


## Example

This Visual Basic for Applications (VBA) example uses the  **[GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md)** method to resolve the **[Recipient](Outlook.Recipient.md)** object representing Dan Wilson, and then returns Dan's shared default **Calendar** folder. To run this example, replace 'Dan Wilson' with a valid recipient name and make sure the calendar is shared and you have permissions to view the calendar.


```vb
Sub ResolveName() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myRecipient As Outlook.Recipient 
 
 Dim CalendarFolder As Outlook.Folder 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myRecipient = myNamespace.CreateRecipient("Dan Wilson") 
 
 myRecipient.Resolve 
 
 If myRecipient.Resolved Then 
 
 Call ShowCalendar(myNamespace, myRecipient) 
 
 End If 
 
End Sub 
 
 
 
Sub ShowCalendar(myNamespace, myRecipient) 
 
 Dim CalendarFolder As Folder 
 
 
 
 Set CalendarFolder = _ 
 
 myNamespace.GetSharedDefaultFolder _ 
 
 (myRecipient, olFolderCalendar) 
 
 CalendarFolder.Display 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]