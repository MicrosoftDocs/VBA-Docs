---
title: Permission members (Office)
ms.prod: office
ms.assetid: 75614d24-cd47-ef9b-aba5-112206daa358
ms.date: 01/30/2019
localization_priority: Normal
---


# Permission members (Office)

The **Permission** property of the **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **Permission** object.


## Methods

|Name|Description|
|:-----|:-----|
|[Add](../../Office.Permission.Add.md)|Creates a set of permissions on the active document for the specified user. Returns a **UserPermission** object.|
|[ApplyPolicy](../../Office.Permission.ApplyPolicy.md)|Applies the specified permission policy to the active document.|
|[RemoveAll](../../Office.Permission.RemoveAll.md)|Removes all **UserPermission** objects from the **Permission** collection of the active document.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.Permission.Application.md)|Gets an **Application** object that represents the container application for the **Permission** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Count](../../Office.Permission.Count.md)|Gets a **Long** indicating the number of items in the **Permission** object. Read-only.|
|[Creator](../../Office.Permission.Creator.md)|Gets a 32-bit integer that indicates the application in which the **Permission** object was created. Read-only.|
|[DocumentAuthor](../../Office.Permission.DocumentAuthor.md)|Gets or sets the name in email form of the author of the active document. Read/write.|
|[Enabled](../../Office.Permission.Enabled.md)|Gets or sets a **Boolean** value that indicates whether permissions are enabled on the active document. Read/write.|
|[EnableTrustedBrowser](../../Office.Permission.EnableTrustedBrowser.md)|Gets or sets a value indicating whether to enable a browser from a trusted source. Read/write.|
|[Item](../../Office.Permission.Item.md)|Gets a **UserPermission** object that is a member of the **Permission** collection. The **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. Read-only.|
|[Parent](../../Office.Permission.Parent.md)|Gets the **Parent** object for the **Permission** object. Read-only.|
|[PermissionFromPolicy](../../Office.Permission.PermissionFromPolicy.md)|Gets a **Boolean** value that indicates whether a permission policy has been applied to the active document. Read-only.|
|[PolicyDescription](../../Office.Permission.PolicyDescription.md)|Gets the description of the permissions policy applied to the active document. Read-only.|
|[PolicyName](../../Office.Permission.PolicyName.md)|Gets the name of the permissions policy applied to the active document. Read-only.|
|[RequestPermissionURL](../../Office.Permission.RequestPermissionURL.md)|Gets or sets the file or Web site URL to visit or the email address to contact for users who need additional permissions on the active document. Read/write.|
|[StoreLicenses](../../Office.Permission.StoreLicenses.md)|Gets or sets a **Boolean** value that indicates whether the user's license to view the active document should be cached to allow offline viewing when the user cannot connect to a rights management server. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]