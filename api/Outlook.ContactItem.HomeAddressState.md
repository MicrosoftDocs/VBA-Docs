---
title: ContactItem.HomeAddressState property (Outlook)
keywords: vbaol11.chm1017
f1_keywords:
- vbaol11.chm1017
ms.prod: outlook
api_name:
- Outlook.ContactItem.HomeAddressState
ms.assetid: bc052902-1e38-3d6a-1b7b-308861357731
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.HomeAddressState property (Outlook)

Returns or sets a  **String** representing the state portion of the home address for the contact. Read/write.


## Syntax

_expression_. `HomeAddressState`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

This property is parsed from the  **[HomeAddress](Outlook.ContactItem.HomeAddress.md)** property, but may be changed or entered independently should it be parsed incorrectly. Note that any such changes or entries to this property will be overwritten by any subsequent changes or entries to **HomeAddress**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]