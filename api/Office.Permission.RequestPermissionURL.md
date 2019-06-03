---
title: Permission.RequestPermissionURL property (Office)
keywords: vbaof11.chm261009
f1_keywords:
- vbaof11.chm261009
ms.prod: office
api_name:
- Office.Permission.RequestPermissionURL
ms.assetid: 7d37d706-a7bf-9cb0-8930-299bd2bf37b0
ms.date: 01/22/2019
localization_priority: Normal
---


# Permission.RequestPermissionURL property (Office)

Gets or sets the file or website URL to visit or the email address to contact for users who need additional permissions on the active document. Read/write.


## Syntax

_expression_.**RequestPermissionURL**

_expression_ A variable that represents a **[Permission](Office.Permission.md)** object.


## Remarks

The **RequestPermissionURL** setting corresponds to the **Users can request additional permissions from** option in the permissions user interface. Use the **RequestPermissionURL** property to specify a file, a website, or an email contact from which users can request, or learn how to request, additional permissions on the active document; for example:

- A web address: `https://companyserver/request_permissions.asp`
    
- A file: `\\companyserver\share\requesting_permissions.txt`
    
- An email address: `mailto:permissionsmgr@example.com?Subject=Request%20permissions`
    

## Example

The following example displays information about the permissions settings of the active document, including the **RequestPermissionURL** setting.


```vb
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." & vbCrLf 
 strIRMInfo = strIRMInfo & " View in trusted browser: " & _ 
 irmPermission.EnableTrustedBrowser & vbCrLf & _ 
 " Document author: " & irmPermission.DocumentAuthor & vbCrLf & _ 
 " Users with permissions: " & irmPermission.Count & vbCrLf & _ 
 " Cache licenses locally: " & irmPermission.StoreLicenses & vbCrLf & _ 
 " Request permission URL: " & irmPermission.RequestPermissionURL & vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo & " Permissions applied from policy:" & vbCrLf & _ 
 " Policy name: " & irmPermission.PolicyName & vbCrLf & _ 
 " Policy description: " & irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo & " Default permissions applied." 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also

- [Permission object members](overview/library-reference/permission-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]