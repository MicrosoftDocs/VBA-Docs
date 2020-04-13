---
title: ContactItem.HomeAddressPostalCode property (Outlook)
keywords: vbaol11.chm1015
f1_keywords:
- vbaol11.chm1015
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressPostalCode
ms.assetid: 28d65f71-6be6-5d9e-0935-7f09a5f9fa94
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressPostalCode property (Outlook)

Returns or sets a **String** representing the postal code portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressPostalCode`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]