---
title: SharingItem.AddBusinessCard method (Outlook)
keywords: vbaol11.chm3217
f1_keywords:
- vbaol11.chm3217
ms.prod: outlook
api_name:
- Outlook.SharingItem.AddBusinessCard
ms.assetid: fa3fa071-b43c-c2d1-7d7c-dc52ab9a1681
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.AddBusinessCard method (Outlook)

Appends contact information based on the Electronic Business Card (EBC) associated with the specified  **[ContactItem](Outlook.ContactItem.md)** object to the **[SharingItem](Outlook.SharingItem.md)** object.


## Syntax

_expression_. `AddBusinessCard`( `_contact_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _contact_|Required| **ContactItem**|The contact item from which to obtain the business card information.|

## Remarks

This method adds contact information, generated from the information stored in the  **ContactItem** object, to the existing **SharingItem** object. The information included depends on the value of the **[BodyFormat](Outlook.SharingItem.BodyFormat.md)** property for the **SharingItem** object:



| **Property value**| **Result**|
| **olFormatPlain**|A vCard (.vcf) file is created and added to the  **[Attachments](Outlook.Attachments.md)** collection of the **SharingItem** object.|
| **olFormatRichText**|A vCard (.vcf) file is created and added to the  **Attachments** collection of the **SharingItem** object.|
| **olFormatHTML**|An image of the business card is generated and included in the  **[Body](Outlook.MailItem.Body.md)** property of the **SharingItem** object, and a vCard (.vcf) file is created and added to the **[Attachments](Outlook.Attachments.md)** collection of the **SharingItem** object.|

> [!NOTE] 
> The attached vCard file contains only the contact information included in the Electronic Business Card associated with the  **ContactItem** object. Any contact information not displayed in the Electronic Business Card is excluded from the vCard file.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]