---
title: MsoEnvelope.Item property (Office)
keywords: vbaof11.chm11003
f1_keywords:
- vbaof11.chm11003
ms.prod: office
api_name:
- Office.MsoEnvelope.Item
ms.assetid: cc13343c-dea5-152f-b123-441a4120c22c
ms.date: 01/22/2019
localization_priority: Normal
---


# MsoEnvelope.Item property (Office)

Gets a **MailItem** object that can be used to send the document as an email. Read-only.


## Syntax

_expression_.**Item**

_expression_ Required. A variable that represents an **[MsoEnvelope](Office.MsoEnvelope.md)** object.


## Example

The following example sends the active Microsoft Word document as an email to the email address that you pass to the subroutine.


```vb
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the msoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the email message. 
 .Introduction = "Please read this and send me your comments." 
 
 'Return a MailItem object that you can use to send the document. 
 With .Item 
 
 'All of the mail item settings are saved with the document. 
 'When you add a recipient to the Recipients collection 
 'or change other properties these settings will persist. 
 
 .Recipients.Add strRecipient 
 .Subject = "Here is the document." 
 
 'The body of this message will be 
 'the content of the active document. 
 .Send 
 End With 
 End With 
End Sub
```


## See also

- [MsoEnvelope object members](overview/library-reference/msoenvelope-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]


