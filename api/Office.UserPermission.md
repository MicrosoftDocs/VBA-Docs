---
title: UserPermission Object (Office)
keywords: vbaof11.chm260000
f1_keywords:
- vbaof11.chm260000
ms.prod: office
api_name:
- Office.UserPermission
ms.assetid: 24378204-2fdd-47ba-2080-fbc409955325
ms.date: 06/08/2017
---


# UserPermission Object (Office)

Associates a set of permissions on the active document with a single user and an optional expiration date. Represents a member of the active document's  **Permission** collection.


## Remarks

Use the  **Add** method of the **Permission** object to grant specific permissions on the active document to a new user, with an optional expiration date. Use the **Remove** method of the **UserPermission** object to remove a user and the user's permissions.

While some permissions granted through the user interface (such as  **msoPermissionPrint** ) apply to all users, you can use the **UserPermission** object to assign them on a per-user basis with per-user expiration dates.


## Example

The following example determines whether the active document has restricted permissions, then lists users and their assigned permissions by returning the  **UserId**, **Permission**, and **ExpirationDate** properties of each **UserPermission** in the document's **Permission** collection.


```vb
 Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 For Each irmUserPerm In irmPermission 
 strIRMInfo = strIRMInfo &amp; irmUserPerm.UserId &amp; vbCrLf &amp; _ 
 " - Permissions: " &amp; irmUserPerm.Permission &amp; vbCrLf &amp; _ 
 " - Expiration Date: " &amp; irmUserPerm.ExpirationDate &amp; vbCrLf 
 Next 
 MsgBox strIRMInfo, _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Else 
 MsgBox "This document is not restricted.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 End If 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing 

```


## Methods



|**Name**|
|:-----|
|[Remove](Office.UserPermission.Remove.md)|

## Properties



|**Name**|
|:-----|
|[Application](Office.UserPermission.Application.md)|
|[Creator](Office.UserPermission.Creator.md)|
|[ExpirationDate](Office.UserPermission.ExpirationDate.md)|
|[Parent](Office.UserPermission.Parent.md)|
|[Permission](Office.UserPermission.Permission.md)|
|[UserId](Office.UserPermission.UserId.md)|

## See also





[Object Model Reference](./overview/reference-object-library-reference-for-office.md)
