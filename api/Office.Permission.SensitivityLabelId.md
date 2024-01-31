---
title: Permission.SensitivityLabelId property (Office)
description: Gets or sets the sensitivity label id included in user defined protection from Microsoft Purview Information Protection. Read/write.
keywords: vbaof11.chm261009
f1_keywords:
- vbaof11.chm261009
api_name:
- Office.Permission.RequestPermissionURL
ms.assetid: 7d37d706-a7bf-9cb0-8930-299bd2bf37b0
ms.date: 06/10/2022
ms.localizationpriority: medium
---

# Permission.SensitivityLabelId property (Office)

Gets or sets the sensitivity label id included in user defined protection from Microsoft Purview Information Protection. Read/write.

## Syntax

_expression_. **SensitivityLabelId**

_expression_ A variable that represents a **[Permission](Office.Permission.md)**.

## Remarks

The **SensitivityLabelId** setting gets or sets the sensitivity label id included in user defined protection from Microsoft Purview Information Protection. Use the **SensitivityLabelId** property to specify the sensitivity label applied to any user defined permissions. Note that accessing the **SensitivityLabelId** property fails with CRYPT_E_ATTRIBUTES_MISSING if the license does not contain a sensitivity label id. If the license was created from specific permission policy xml (refer to **PermissionFromPolicy**), setting the **SensitivityLabelId** fails with CERTSRV_E_TEMPLATE_DENIED.

## Example

The following example sets and displays the **SensitivityLabelId** setting.

```vb

 Dim irmPermission As Office.Permission 
 Set irmPermission = ActiveWorkbook.Permission 
 If irmPermission.Enabled Then 
 MsgBox irmPermission.SensitivityLabelId, vbInformation + vbOKOnly, " SensitivityLabelId"
 irmPermission. SensitivityLabelId = "24fb0e8b-8044-4bd6-9e86-77f9d21856dc"
 MsgBox irmPermission.SensitivityLabelId, vbInformation + vbOKOnly, " SensitivityLabelId"
End If

 ```
## See also

- [Permission object members](overview/library-reference/permission-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
