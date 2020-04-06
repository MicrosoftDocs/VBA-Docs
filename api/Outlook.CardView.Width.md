---
title: CardView.Width property (Outlook)
keywords: vbaol11.chm2600
f1_keywords:
- vbaol11.chm2600
ms.prod: outlook
api_name:
- Outlook.CardView.Width
ms.assetid: 6140719b-1094-0991-a1d1-8d47e59bd25a
ms.date: 06/08/2017
localization_priority: Normal
---


# CardView.Width property (Outlook)

Returns or sets a  **Long** value indicating the width (in characters) of cards in the **[CardView](Outlook.CardView.md)** object. Read/write.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a [CardView](Outlook.CardView.md) object.


## Remarks

This property can be set to a value between 20 and 1000. If this property is set to a value less than 20, the property is set to 20. If this property is set to a value greater than 1000, the property is set to 1000.

The default value for this property depends on the  **[DefaultItemType](Outlook.Folder.DefaultItemType.md)** property value of the **[Folder](Outlook.Folder.md)** object displayed by the view:



|** **DefaultItemType value****|** **Default value****|
|:-----|:-----|
| **olAppointmentItem**|40|
| **olContactItem**,  **olDistributionListItem**|36|
| **olJournalItem**,  **olMailItem**,  **olNoteItem**,  **olPostItem**|32|
| **olTaskItem**|50|

## See also


[CardView Object](Outlook.CardView.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]