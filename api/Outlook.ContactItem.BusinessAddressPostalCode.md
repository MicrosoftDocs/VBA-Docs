---
title: ContactItem.BusinessAddressPostalCode property (Outlook)
keywords: vbaol11.chm973
f1_keywords:
- vbaol11.chm973
ms.prod: outlook
api_name:
- Outlook.ContactItem.BusinessAddressPostalCode
ms.assetid: 0c9f643a-c29e-4ae5-cea7-f54b3e98b543
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.BusinessAddressPostalCode property (Outlook)

Returns or sets a **String** representing the postal code (ZIP Code) portion of the business address for the contact. Read/write.


## Syntax

_expression_. `BusinessAddressPostalCode`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[BusinessAddress](Outlook.ContactItem.BusinessAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to the **BusinessAddress** property.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]