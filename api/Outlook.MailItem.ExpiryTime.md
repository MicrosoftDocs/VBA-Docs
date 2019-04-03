---
title: MailItem.ExpiryTime property (Outlook)
keywords: vbaol11.chm1334
f1_keywords:
- vbaol11.chm1334
ms.prod: outlook
api_name:
- Outlook.MailItem.ExpiryTime
ms.assetid: 18f6497b-6db5-7ec2-7aa8-ec30531e59ef
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.ExpiryTime property (Outlook)

Returns or sets a  **Date** indicating the date and time at which the item becomes invalid and can be deleted. Read/write.


## Syntax

_expression_. `ExpiryTime`

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Example

This Visual Basic for Applications (VBA) example uses the  **[MailItem.Send](Outlook.MailItem.Send(even).md)** event and sends an item with an automatic expiration date.


```vb
Public WithEvents myItem As MailItem 
 
 
 
Sub SendMyMail() 
 
 Set myItem = Outlook.CreateItem(olMailItem) 
 
 myItem.To = "Laura Jennings" 
 
 myItem.Subject = "Data files information" 
 
 myItem.Send 
 
End Sub 
 
 
 
Private Sub myItem_Send(Cancel As Boolean) 
 
 myItem.ExpiryTime = #2/2/2003 4:00:00 PM# 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]