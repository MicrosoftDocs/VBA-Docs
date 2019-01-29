---
title: UserPermission.Remove method (Office)
keywords: vbaof11.chm260005
f1_keywords:
- vbaof11.chm260005
ms.prod: office
api_name:
- Office.UserPermission.Remove
ms.assetid: d4c8778f-dc1b-7d5b-6a7a-65b91909bfe3
ms.date: 01/29/2019
localization_priority: Normal
---


# UserPermission.Remove method (Office)

Removes the specified **UserPermission** object from the **[Permission](Office.Permission.md)** collection of the active document.


## Syntax

_expression_.**Remove**

_expression_ Required. A variable that represents a **[UserPermission](Office.UserPermission.md)** object.


## Remarks

The **UserPermission** object associates a set of permissions on the active document with a single user and an optional expiration date. The **Remove** method removes the user and the set of user permissions determined by the specified **UserPermission** object.


## Example

The following example removes the second user's permissions on the active document from the document's **Permission** collection.


```vb
 Dim irmPermission As Office.Permission 
 Dim irmUserPerm As Office.UserPermission 
 Set irmPermission = ActiveWorkbook.Permission 
 Set irmUserPerm = irmPermission.Item(2) 
 irmUserPerm.Remove 
 MsgBox "Permission removed.", _ 
 vbInformation + vbOKOnly, "IRM Information" 
 Set irmUserPerm = Nothing 
 Set irmPermission = Nothing 

```


## See also

- [UserPermission object members](overview/Library-Reference/userpermission-members-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]