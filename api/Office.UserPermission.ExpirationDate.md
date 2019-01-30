---
title: UserPermission.ExpirationDate property (Office)
keywords: vbaof11.chm260003
f1_keywords:
- vbaof11.chm260003
ms.prod: office
api_name:
- Office.UserPermission.ExpirationDate
ms.assetid: 769cd094-62c2-a9cd-9214-6fcc799617be
ms.date: 01/29/2019
localization_priority: Normal
---


# UserPermission.ExpirationDate property (Office)

Gets or sets the optional expiration date of the permissions on the active document assigned to the user associated with the specified **UserPermission** object. Read/write.


## Syntax

_expression_.**ExpirationDate**

_expression_ A variable that represents a **[UserPermission](Office.UserPermission.md)** object.


## Return value

Variant


## Remarks

The **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. The **ExpirationDate** property returns or sets the optional expiration date of the specified **UserPermission** object using the local time zone.


## See also

- [UserPermission object members](overview/Library-Reference/userpermission-members-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]