---
title: Permission.Add method (Office)
keywords: vbaof11.chm261004
f1_keywords:
- vbaof11.chm261004
ms.prod: office
api_name:
- Office.Permission.Add
ms.assetid: 9674440f-8b0f-c611-3a02-f0ba1e92be94
ms.date: 06/08/2017
localization_priority: Normal
---


# Permission.Add method (Office)

Creates a set of permissions on the active document for the specified user. Returns a  **UserPermission** object.


## Syntax

_expression_. `Add`( `_UserID_`, `_Permission_`, `_ExpirationDate_` )

 _expression_ Required. A variable that represents a '[Permission](Office.Permission.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _UserID_|Required|**String**|The email address (in the format user@domain.com) of the user to whom permissions on the active document are being granted.|
| _Permission_|Optional|**msoPermission**|The permissions on the active document that are being granted to the specified user.|
| _ExpirationDate_|Optional|**Date**|The expiration date for the permissions that are being granted. **Note: this parameter is not used and will be ignored.**|

## Example

The following example assigns a combination of read and edit permissions on the current document to a user and specifies an expiration date for these document permissions.


```vb
 Dim objUserPerm As Office.UserPermission 
 Set objUserPerm = ActiveWorkbook.Permission.Add( _ 
 "user@domain.com", _ 
 msoPermissionRead + msoPermissionEdit, #12/31/2005#) 
 MsgBox "Permissions added for " &amp; _ 
 objUserPerm.UserId, _ 
 vbInformation + vbOKOnly, _ 
 "Permissions Added" 
 Set objUserPerm = Nothing 

```


## See also


[Permission Object](Office.Permission.md)



[Permission Object Members](overview/Library-Reference/permission-members-office.md)

