---
title: MailItem.Recipients property (Outlook)
keywords: vbaol11.chm1347
f1_keywords:
- vbaol11.chm1347
ms.prod: outlook
api_name:
- Outlook.MailItem.Recipients
ms.assetid: 58897f66-8a6a-e1a9-7e3b-5a84624f899d
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Recipients property (Outlook)

Returns a  **[Recipients](Outlook.Recipients.md)** collection that represents all the recipients for the Outlook item. Read-only.


## Syntax

_expression_. `Recipients`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Remarks

A recipient can be specified by a string representing the recipient's display name, alias, or full SMTP email address.


## Example

This Visual Basic for Applications (VBA) example creates a new email message, uses the  **[Add](Outlook.Recipients.Add.md)** method to add "Dan Wilson" as a **[To](Outlook.MailItem.To.md)** recipient, and displays the message.


```vb
Sub CreateStatusReportToBoss() 
 
 Dim myItem As Outlook.MailItem 
 
 Dim myRecipient As Outlook.Recipient 
 
 
 
 Set myItem = Application.CreateItem(olMailItem) 
 
 Set myRecipient = myItem.Recipients.Add("Dan Wilson") 
 
 myItem.Subject = "Status Report" 
 
 myItem.Display 
 
End Sub
```


## See also

- [Send an email given the SMTP address of an account](../outlook/How-to/Items-Folders-and-Stores/send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
