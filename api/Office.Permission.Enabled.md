---
title: Permission.Enabled property (Office)
keywords: vbaof11.chm261008
f1_keywords:
- vbaof11.chm261008
ms.prod: office
api_name:
- Office.Permission.Enabled
ms.assetid: e77fab6f-0191-3ba4-d418-dc25dc79422d
ms.date: 01/22/2019
localization_priority: Normal
---


# Permission.Enabled property (Office)

Gets or sets a **Boolean** value that indicates whether permissions are enabled on the active document. Read/write.


## Syntax

_expression_.**Enabled**

_expression_ Required. A variable that represents a **[Permission](Office.Permission.md)** object.


## Remarks

Use the **Enabled** property to determine whether permissions are restricted on the active document, and to enable or disable permissions. Set **Enabled** to **False** to disable permissions and to remove all users, other than the document author, and their permissions.

When permissions are disabled, the **Count** property of the **Permission** object returns 0 (zero); however, when permissions are re-enabled, the permissions of the document author remain intact.


## See also

- [Permission object members](overview/library-reference/permission-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]