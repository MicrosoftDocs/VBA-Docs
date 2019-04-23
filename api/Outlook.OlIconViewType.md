---
title: OlIconViewType enumeration (Outlook)
keywords: vbaol11.chm3124
f1_keywords:
- vbaol11.chm3124
ms.prod: outlook
api_name:
- Outlook.OlIconViewType
ms.assetid: 81fefee1-25b3-dbe5-0d10-047259e273a6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlIconViewType enumeration (Outlook)

Indicates the view mode when using an  **[IconView](Outlook.IconView.md)** object.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olIconViewLarge**|0|Displays Outlook items as large icons, with the description for the Outlook item below the icon.|
| **olIconViewList**|2|Displays Outlook items as a list of small icons, with the description for the Outlook item next to the icon.|
| **olIconViewSmall**|1|Displays Outlook items as a collection of small icons, with the description for the Outlook item next to the icon.|

## Remarks

For  **olIconViewLarge** and **olIconViewSmall**, the actual placement and orientation of icons displayed in the [IconView](Outlook.IconView.md) object is determined by the [OlIconViewPlacement](Outlook.OlIconViewPlacement.md) constant specified in the [IconPlacement](Outlook.IconView.IconPlacement.md) property of the **IconView** object. The **IconPlacement** property value is ignored when **olIconViewList** is selected.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]