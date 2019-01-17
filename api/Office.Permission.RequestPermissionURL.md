---
title: Permission.RequestPermissionURL property (Office)
keywords: vbaof11.chm261009
f1_keywords:
- vbaof11.chm261009
ms.prod: office
api_name:
- Office.Permission.RequestPermissionURL
ms.assetid: 7d37d706-a7bf-9cb0-8930-299bd2bf37b0
ms.date: 06/08/2017
localization_priority: Normal
---


# Permission.RequestPermissionURL property (Office)

Gets or sets the file or Web site URL to visit or the email address to contact for users who need additional permissions on the active document. Read/write.


## Syntax

_expression_. `RequestPermissionURL`

_expression_ A variable that represents a [Permission](Office.Permission.md) object.


## Remarks

The ** RequestPermissionURL** setting corresponds to the **Users can request additional permissions from** option in the permissions user interface. Use the **RequestPermissionURL** property to specify a file, a Web site, or an email contact from which users can request, or learn how to request, additional permissions on the active document, for example:


- A Web address:  `https://companyserver/request_permissions.asp`
    
- A file:  `\\companyserver\share\requesting_permissions.txt`
    
- An email address:  `mailto:permissionsmgr@example.com?Subject=Request%20permissions`
    

## Example

The following example displays information about the permissions settings of the active document, including the  **RequestPermissionURL** setting.


```vb
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." &amp; vbCrLf 
 strIRMInfo = strIRMInfo &amp; " View in trusted browser: " &amp; _ 
 irmPermission.EnableTrustedBrowser &amp; vbCrLf &amp; _ 
 " Document author: " &amp; irmPermission.DocumentAuthor &amp; vbCrLf &amp; _ 
 " Users with permissions: " &amp; irmPermission.Count &amp; vbCrLf &amp; _ 
 " Cache licenses locally: " &amp; irmPermission.StoreLicenses &amp; vbCrLf &amp; _ 
 " Request permission URL: " &amp; irmPermission.RequestPermissionURL &amp; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &amp; " Permissions applied from policy:" &amp; vbCrLf &amp; _ 
 " Policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Policy description: " &amp; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &amp; " Default permissions applied." 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also


[Permission Object](Office.Permission.md)



[Permission Object Members](./overview/Library-Reference/permission-members-office.md)

