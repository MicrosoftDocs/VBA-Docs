---
title: SharingItem.HTMLBody property (Outlook)
keywords: vbaol11.chm641
f1_keywords:
- vbaol11.chm641
ms.prod: outlook
api_name:
- Outlook.SharingItem.HTMLBody
ms.assetid: cd181b3f-e990-3d41-aa30-ec51361c605d
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.HTMLBody property (Outlook)

Returns or sets a **String** representing the HTML body of the specified **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_. `HTMLBody`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The **HTMLBody** property should be an HTML syntax string.

Setting the  **HTMLBody** property sets the **[EditorType](Outlook.Inspector.EditorType.md)** property of the item's **[Inspector](Outlook.Inspector.md)** to **olEditorHTML**.

Setting the  **HTMLBody** property will always update the **[Body](Outlook.SharingItem.Body.md)** property immediately.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]