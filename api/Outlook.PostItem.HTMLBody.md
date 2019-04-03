---
title: PostItem.HTMLBody property (Outlook)
keywords: vbaol11.chm1548
f1_keywords:
- vbaol11.chm1548
ms.prod: outlook
api_name:
- Outlook.PostItem.HTMLBody
ms.assetid: 5db93b3c-96b0-ce14-4d53-cbc113c2c14c
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.HTMLBody property (Outlook)

Returns or sets a  **String** representing the HTML body of the specified item. Read/write.


## Syntax

_expression_. `HTMLBody`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

The  **HTMLBody** property should be an HTML syntax string.

Setting the  **HTMLBody** property sets the **[EditorType](Outlook.Inspector.EditorType.md)** property of the item's **[Inspector](Outlook.Inspector.md)** to **olEditorHTML**.

Setting the  **HTMLBody** property will always update the **[Body](Outlook.PostItem.Body.md)** property immediately.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]