---
title: ContactItem.HomeAddressPostOfficeBox property (Outlook)
keywords: vbaol11.chm1016
f1_keywords:
- vbaol11.chm1016
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressPostOfficeBox
ms.assetid: 9c1b310d-13d8-407c-a97e-a52405e37fb2
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressPostOfficeBox property (Outlook)

Returns or sets a **String** the post office box number portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressPostOfficeBox`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]