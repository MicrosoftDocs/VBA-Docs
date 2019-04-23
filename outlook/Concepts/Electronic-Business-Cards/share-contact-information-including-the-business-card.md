---
title: Share Contact Information Including the Business Card
ms.prod: outlook
ms.assetid: 57218e2f-a6fd-bd52-0065-b8ff8b480d3c
ms.date: 06/08/2017
localization_priority: Normal
---


# Share Contact Information Including the Business Card

You can use the  **[ForwardAsVcard](../../../api/Outlook.ContactItem.ForwardAsVcard.md)** and **[ForwardAsBusinessCard](../../../api/Outlook.ContactItem.ForwardAsBusinessCard.md)** method of the **[ContactItem](../../../api/Outlook.ContactItem.md)** object to create a new **[MailItem](../../../api/Outlook.MailItem.md)** object that contains the contact information from the specified **ContactItem** attached as a vCard (.vcf) file, or you can use the **[AddBusinessCard](../../../api/Outlook.MailItem.AddBusinessCard.md)** method of the **MailItem** object to attach the contact information for a specified **ContactItem** as a vCard file. If you use the **ForwardAsBusinessCard** or **AddBusinessCard** methods, an image of the business card is also appended to the body of the mail item if the **[BodyFormat](../../../api/Outlook.MailItem.BodyFormat.md)** property of the **MailItem** object is set to **olFormatHTML**.

The following code sample in Microsoft Visual Basic for Applications (VBA) is a function,  `ForwardContactItem`, that accepts a  **ContactItem** object as a parameter and forwards the **ContactItem** object as an attachment to a new mail item. `ForwardContactItem` first checks if the object is a valid object. If the object is valid, `ForwardContactItem` calls the **ForwardAsBusinessCard** method of the **ContactItem** object to create a new **MailItem** object that has the contact information attached as a vCard. `ForwardContactItem` then displays and returns the **MailItem** object.



```vb
Private Function ForwardContactItem(objContactItem As Outlook.ContactItem) As Outlook.MailItem 
 
 Dim objMailItem As MailItem 
 
 On Error GoTo ErrRoutine 
 
 If objContactItem Is Nothing Then 
 ForwardContactItem = Nothing 
 Else 
 ' Forward the contact item, including a business card 
 ' image, and display the new MailItem object. 
 Set objMailItem = objContactItem.ForwardAsBusinessCard 
 objMailItem.Display 
 ForwardContactItem = objMailItem 
 End If 
 
EndRoutine: 
 Exit Function 
 
ErrRoutine: 
 MsgBox Err.Number & " - " & Err.Description, _ 
 vbOKOnly Or vbCritical, _ 
 "ForwardContactItem" 
 GoTo EndRoutine 
End Function
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]