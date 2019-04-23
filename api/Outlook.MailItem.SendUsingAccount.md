---
title: MailItem.SendUsingAccount property (Outlook)
keywords: vbaol11.chm1390
f1_keywords:
- vbaol11.chm1390
ms.prod: outlook
api_name:
- Outlook.MailItem.SendUsingAccount
ms.assetid: d4e49128-a63a-d761-90b9-9e1a3305adc7
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.SendUsingAccount property (Outlook)

Returns or sets an **[Account](Outlook.Account.md)** object that represents the account under which the **[MailItem](Outlook.MailItem.md)** is to be sent. Read/write.


## Syntax

_expression_. SendUsingAccount

_expression_ An expression that returns a [MailItem](Outlook.MailItem.md) object.


## Remarks

The **SendUsingAccount** property can be used to specify the account that should be used to send the **MailItem** when the **[Send](Outlook.MailItem.Send(method).md)** method is called. This property returns **Null** (**Nothing** in Visual Basic) if the account specified for the **MailItem** no longer exists.


## Example

The following code sample in Microsoft Visual Basic for Applications enumerates the **[Accounts](Outlook.Accounts.md)** collection to find a Pop3 account. If the account is found, a message is created programmatically, and the **SendUsingAccount** property is assigned to the Pop3 account. Note that you must assign the **SendUsingAccount** property before you call the **Send** method.


```vb
Sub SendUsingAccount() 
 
 Dim oAccount As Outlook.account 
 
 For Each oAccount In Application.Session.Accounts 
 
 If oAccount.AccountType = olPop3 Then 
 
 Dim oMail As Outlook.MailItem 
 
 Set oMail = Application.CreateItem(olMailItem) 
 
     oMail.Subject = "Sent using POP3 Account" 
 
     oMail.Recipients.Add ("someone@example.com") 
 
     oMail.Recipients.ResolveAll 
 
 Set oMail.SendUsingAccount = oAccount 
 
     oMail.Send 
 
 End If 
 
 Next 
 
End Sub
```


## See also

- [Send an email given the SMTP address of an account](../outlook/How-to/Items-Folders-and-Stores/send-an-e-mail-given-the-smtp-address-of-an-account-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
