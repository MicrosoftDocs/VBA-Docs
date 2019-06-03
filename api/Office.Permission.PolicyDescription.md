---
title: Permission.PolicyDescription property (Office)
keywords: vbaof11.chm261011
f1_keywords:
- vbaof11.chm261011
ms.prod: office
api_name:
- Office.Permission.PolicyDescription
ms.assetid: 1ca10f9f-f03a-3a3f-2c12-21831a092f23
ms.date: 01/22/2019
localization_priority: Normal
---


# Permission.PolicyDescription property (Office)

Gets the description of the permissions policy applied to the active document. Read-only.


## Syntax

_expression_.**PolicyDescription**

_expression_ A variable that represents a **[Permission](Office.Permission.md)** object.


## Remarks

Microsoft Office Information Rights Management supports the use of administrative permission policies that list users and groups and their document permissions. The **PolicyDescription** property returns the description of the policy applied to the active document, or a default value if a policy was not used.


## Example

The following example displays permission policy information about the active document.


```vb
 Dim irmPermission As Office.Permission 
 Dim strIRMInfo As String 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 strIRMInfo = "Permissions are restricted on this document." & vbCrLf 
 If irmPermission.PermissionFromPolicy Then 
 strIRMInfo = strIRMInfo & " Permissions applied from policy:" & vbCrLf & _ 
 " Policy name: " & irmPermission.PolicyName & vbCrLf & _ 
 " Policy description: " & irmPermission.PolicyDescription 
 Else 
 strIRMInfo = strIRMInfo & " Default permissions applied." & vbCrLf & _ 
 " Default policy name: " & irmPermission.PolicyName & vbCrLf & _ 
 " Default policy description: " & irmPermission.PolicyDescription 
 End If 
 Else 
 strIRMInfo = "Permission are NOT restricted on this document." 
 End If 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also

- [Permission object members](overview/library-reference/permission-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]