---
title: MsoEnvelope object (Office)
keywords: vbaof11.chm245000
f1_keywords:
- vbaof11.chm245000
ms.prod: office
api_name:
- Office.MsoEnvelope
ms.assetid: 64cfde6b-cd71-1d7b-0e8f-1181d88d9457
ms.date: 01/22/2019
localization_priority: Normal
---


# MsoEnvelope object (Office)

Provides access to functionality that lets you send documents as email messages directly from Microsoft Office applications.


## Remarks

Use the **MailEnvelope** property of the **Document** object, **Chart** object, or **Worksheet** object (depending on the application you are using) to return an **MsoEnvelope** object.


## Example

The following example sends the active Microsoft Word document as an email message to the email address that you pass to the subroutine.


```vb
Sub SendMail(ByVal strRecipient As String) 
 
 'Use a With...End With block to reference the MsoEnvelope object. 
 With Application.ActiveDocument.MailEnvelope 
 
 'Add some introductory text before the body of the email. 
 .Introduction = "Please read this and send me your comments." 
 
 'Return a Microsoft Outlook MailItem object that 
 'you can use to send the document. 
 With .Item 
 
 'All of the mail item settings are saved with the document. 
 'When you add a recipient to the Recipients collection 
 'or change other properties, these settings persist. 
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
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]


