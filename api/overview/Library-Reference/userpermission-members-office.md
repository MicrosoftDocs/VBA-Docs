---
title: UserPermission members (Office)
description: Associates a set of permissions on the active document with a single user and an optional expiration date.
ms.prod: office
ms.assetid: b9fdae9a-719b-9e1d-42aa-7553de91f9d1
ms.date: 01/30/2019
localization_priority: Normal
---


# UserPermission members (Office)

Associates a set of permissions on the active document with a single user and an optional expiration date. Represents a member of the active document's **Permission** collection.


## Methods

|Name|Description|
|:-----|:-----|
|[Remove](../../Office.UserPermission.Remove.md)|Removes the specified **UserPermission** object from the **[Permission](../../Office.Permission.md)** collection of the active document.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.UserPermission.Application.md)|Gets an **Application** object that represents the container application for the **UserPermission** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](../../Office.UserPermission.Creator.md)|Gets a 32-bit integer that indicates the application in which the **UserPermission** object was created. Read-only.|
|[ExpirationDate](../../Office.UserPermission.ExpirationDate.md)|Gets or sets the optional expiration date of the permissions on the active document assigned to the user associated with the specified **UserPermission** object. Read/write.|
|[Parent](../../Office.UserPermission.Parent.md)|Gets the **Parent** object for the **UserPermission** object. Read-only.|
|[Permission](../../Office.UserPermission.Permission.md)| Returns or sets a **MsoPermission** constant as a **Long** value representing the permissions on the active document assigned to the user associated with the specified **UserPermission** object. Read/write.|
|[UserId](../../Office.UserPermission.UserId.md)|Gets the email name of the user whose permissions on the active document are determined by the specified **[UserPermission](../../Office.UserPermission.md)** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]