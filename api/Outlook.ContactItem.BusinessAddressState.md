---
title: ContactItem.BusinessAddressState property (Outlook)
keywords: vbaol11.chm975
f1_keywords:
- vbaol11.chm975
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressState
ms.assetid: 0d8d9136-6d41-b0ed-f320-6e26fca15cf7
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.BusinessAddressState property (Outlook)

Returns or sets a  **String** representing the state code portion of the business address for the contact. Read/write.


## Syntax

_expression_. `BusinessAddressState`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[BusinessAddress](Outlook.ContactItem.BusinessAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]