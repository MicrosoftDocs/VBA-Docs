---
title: ContactItem.MailingAddressState property (Outlook)
keywords: vbaol11.chm1039
f1_keywords:
- vbaol11.chm1039
ms.prod: outlook
api_name:
- Outlook.ContactItem.MailingAddressState
ms.assetid: 9e15bba8-2256-fd1a-60ae-ac63d6d4f4e3
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MailingAddressState property (Outlook)

Returns or sets a  **String** representing the state code portion for the selected mailing address of the contact. Read/write.


## Syntax

_expression_. `MailingAddressState`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property replicates the property indicated by the  **[SelectedMailingAddress](Outlook.ContactItem.SelectedMailingAddress.md)** property, which is one of the following **OlMailingAddress** constants: **olBusiness**, **olHome**, **olNone**, or **olOther**. While it can be changed or entered independently, any such changes or entries to this property will be overwritten by any subsequent changes or entries to the property indicated by **SelectedMailingAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]