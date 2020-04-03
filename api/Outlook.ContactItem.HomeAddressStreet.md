---
title: ContactItem.HomeAddressStreet property (Outlook)
keywords: vbaol11.chm1018
f1_keywords:
- vbaol11.chm1018
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressStreet
ms.assetid: 9a7af500-e817-6fb1-89b4-6b0ef70741bf
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressStreet property (Outlook)

Returns or sets a  **String** representing the street portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressStreet`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]