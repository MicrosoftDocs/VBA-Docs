---
title: Permission.ApplyPolicy method (Office)
keywords: vbaof11.chm261005
f1_keywords:
- vbaof11.chm261005
ms.prod: office
api_name:
- Office.Permission.ApplyPolicy
ms.assetid: d1904d11-d212-de2f-19cb-78911136ccd7
ms.date: 01/22/2019
localization_priority: Normal
---


# Permission.ApplyPolicy method (Office)

Applies the specified permission policy to the active document.


## Syntax

_expression_.**ApplyPolicy** (_FileName_)

_expression_ A variable that represents a **[Permission](Office.Permission.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**| The path and file name of the permission policy template file.|

## Remarks

Microsoft Office Information Rights Management supports the use of administrative permission policies that list users and groups and their document permissions. The **ApplyPolicy** method applies a permission policy to the active document.


## Example

The following example enables permissions on the active document and applies an administrative permission policy.


```vb
 Dim irmPermission As Office.Permission 
 Set irmPermission = ActiveWorkbook.Permission 
 Dim strIRMInfo As String 
 Select Case irmPermission.Enabled 
 Case True 
 strIRMInfo = "Permissions are already restricted on this document." 
 Case False 
 With irmPermission 
 .Enabled = True 
 .ApplyPolicy ("\\server\share\permissionpolicy.xml") 
 End With 
 strIRMInfo = "Permissions are now restricted on this document " & _ 
 vbCrLf & _ 
 " and the permission policy has been applied." 
 End Select 
 MsgBox strIRMInfo, vbInformation + vbOKOnly, "IRM Information" 
 Set irmPermission = Nothing 

```


## See also

- [Permission object members](overview/library-reference/permission-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]