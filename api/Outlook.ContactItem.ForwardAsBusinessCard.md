---
title: ContactItem.ForwardAsBusinessCard method (Outlook)
keywords: vbaol11.chm1094
f1_keywords:
- vbaol11.chm1094
ms.prod: outlook
api_name:
- Outlook.ContactItem.ForwardAsBusinessCard
ms.assetid: 2f1a74c3-86f0-a054-75e2-272dbb261fb7
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.ForwardAsBusinessCard method (Outlook)

Creates a new **[MailItem](Outlook.MailItem.md)** object containing contact information and, optionally, an Electronic Business Card (EBC) image based on the specified **[ContactItem](Outlook.ContactItem.md)** object.


## Syntax

_expression_. `ForwardAsBusinessCard`

 _expression_ An expression that returns a [ContactItem](Outlook.ContactItem.md) object.


## Return value

A **MailItem** object that represents the new email item containing the business card information.


## Remarks

This method creates a new Outlook mail item based on the information stored in the **ContactItem** object. The information included in the Outlook mail item depends on the value of the **[BodyFormat](Outlook.MailItem.BodyFormat.md)** property for the **MailItem** object:



| **Property value**| **Result**|
|---|---|
| **olFormatPlain**|A vCard (.vcf) file is created and added to the **[Attachments](Outlook.Attachments.md)** collection of the **MailItem** object.|
| **olFormatRichText**|A vCard file is created and added to the **Attachments** collection of the **MailItem** object.|
| **olFormatHTML**|An image of the Electronic Business Card is generated and included in the **[Body](Outlook.MailItem.Body.md)** property of the **MailItem** object, and a vCard file is created and added to the **[Attachments](Outlook.Attachments.md)** collection of the **MailItem** object.|

> [!NOTE] 
> The attached vCard file contains only the contact information included in the Electronic Business Card associated with the **ContactItem** object. Any contact information not displayed in the Electronic Business Card is excluded from the vCard file.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
