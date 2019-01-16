---
title: Permission object (Office)
keywords: vbaof11.chm261000
f1_keywords:
- vbaof11.chm261000
ms.prod: office
api_name:
- Office.Permission
ms.assetid: 4bdf7058-d4ba-0bd4-c5cd-141d67245ced
ms.date: 06/08/2017
localization_priority: Normal
---


# Permission object (Office)

The  **Permission** property of the **Document** object in Microsoft Word, a **Workbook** object in Microsoft Excel, and a **Presentation** object in Microsoft PowerPoint returns a **Permission** object.


## Remarks

Use the  **Permission** object to restrict permissions to the active document and to return or set specific permissions settings.

The  **Permission** object gives access to a collection of **UserPermission** objects. Use the **UserPermission** object to associate specific sets of rights with individual users. While some permissions granted through the user interface (such as **msoPermissionPrint**) apply to all users, you can use the **UserPermission** object to assign them on a per-user basis with per-user expiration dates.

Microsoft Office Information Rights Management supports the use of administrative permission policies which list users and groups and their document permissions. Use the  **ApplyPolicy** method to apply a permission policy, and the **PermissionFromPolicy**, **PolicyName**, and **PolicyDescription** properties to return policy information.

The  **Permission** object model is available whether permissions are restricted on the active document or not . The **Permission** property of the **Document**, **Workbook**, and **Presentation** objects does not return **Nothing** when the active document does not have restricted permissions. Use the **Enabled** property to determine whether a document has restricted permissions.

Use of the  **Permission** object raises an error when the Windows Rights Management client is not installed.


## Example

The following example returns information about the permissions settings on the active document.


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
 " Cache licenses: " &amp; irmPermission.StoreLicenses &amp; vbCrLf &amp; _ 
 " Request permission URL: " &amp; irmPermission.RequestPermissionURL &amp; vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo &amp; " Permissions applied from policy:" &amp; vbCrLf &amp; _ 
 " Policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Policy description: " &amp; irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo &amp; " Default permissions applied." &amp; vbCrLf &amp; _ 
 " Default policy name: " &amp; irmPermission.PolicyName &amp; vbCrLf &amp; _ 
 " Default policy description: " &amp; irmPermission.PolicyDescription 
 End If 
 Else 
 strIRMInfo = "Permissions are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing
```


## Methods



|Name|
|:-----|
|[Add](Office.Permission.Add.md)|
|[ApplyPolicy](Office.Permission.ApplyPolicy.md)|
|[RemoveAll](Office.Permission.RemoveAll.md)|

## Properties



|Name|
|:-----|
|[Application](Office.Permission.Application.md)|
|[Count](Office.Permission.Count.md)|
|[Creator](Office.Permission.Creator.md)|
|[DocumentAuthor](Office.Permission.DocumentAuthor.md)|
|[Enabled](Office.Permission.Enabled.md)|
|[EnableTrustedBrowser](Office.Permission.EnableTrustedBrowser.md)|
|[Item](Office.Permission.Item.md)|
|[Parent](Office.Permission.Parent.md)|
|[PermissionFromPolicy](Office.Permission.PermissionFromPolicy.md)|
|[PolicyDescription](Office.Permission.PolicyDescription.md)|
|[PolicyName](Office.Permission.PolicyName.md)|
|[RequestPermissionURL](Office.Permission.RequestPermissionURL.md)|
|[StoreLicenses](Office.Permission.StoreLicenses.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)
