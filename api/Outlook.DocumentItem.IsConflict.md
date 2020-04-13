---
title: DocumentItem.IsConflict property (Outlook)
keywords: vbaol11.chm1222
f1_keywords:
- vbaol11.chm1222
ms.prod: outlook
api_name:
- Outlook.DocumentItem.IsConflict
ms.assetid: 63d799ea-ceb6-b070-16a6-629ee3ef2346
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.IsConflict property (Outlook)

Returns a **Boolean** that determines if the item is in conflict. Read-only.


## Syntax

_expression_. `IsConflict`

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Remarks

Whether or not an item is in conflict is determined by the state of the application. For example, when a user is offline and tries to access an online folder the action will fail. In this scenario, the  **IsConflict** property will return **True**.

If  **True**, the specified item is in conflict.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]