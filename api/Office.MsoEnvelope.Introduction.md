---
title: MsoEnvelope.Introduction property (Office)
keywords: vbaof11.chm11001
f1_keywords:
- vbaof11.chm11001
ms.prod: office
api_name:
- Office.MsoEnvelope.Introduction
ms.assetid: f37129d4-2a68-1623-272b-f71dfdeec59b
ms.date: 01/22/2019
localization_priority: Normal
---


# MsoEnvelope.Introduction property (Office)

Sets or gets the introductory text that is included with a document that is sent by using the **MsoEnvelope** object. The introductory text is included at the top of the document in the email. Read/write.


## Syntax

_expression_.**Introduction**

_expression_ A variable that represents an **[MsoEnvelope](Office.MsoEnvelope.md)** object.


## Example

The following example sends the active Microsoft Word document as an email to the email address that you pass to the subroutine.


```vb
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the MsoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the email. 
 .Introduction = "Please read this and send me your comments." 
 
 'Return a MailItem object that you can use to send the document. 
 With .Item 
 
 'All of the mail item settings are saved with the document. 
 'When you add a recipient to the Recipients collection 
 'or change other properties, these settings will persist. 
 
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


