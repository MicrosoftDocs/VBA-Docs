---
title: Application.OptionsSecurityEx method (Project)
keywords: vbapj.chm652
f1_keywords:
- vbapj.chm652
ms.prod: project-server
api_name:
- Project.Application.OptionsSecurityEx
ms.assetid: 9c6e0c77-6873-1a90-fb85-ca33ca7c9ec1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OptionsSecurityEx method (Project)

Sets legacy security options that are available in the **Trust Center** dialog box.


## Syntax

_expression_. `OptionsSecurityEx`( `_RemoveFileProperties_`, `_TrustWSS_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RemoveFileProperties_|Optional|**Boolean**|**True** if Project removes personal information from file properties upon saving. The default value is **False**. Corresponds to the **Document-specific settings** section on the **Privacy Options** tab of the **Trust Center** dialog box.|
| _TrustWSS_|Optional|**Boolean**|**True** if Project Server and project workspace sites need not be added to the Internet Explorer Trusted Sites list. **False** if the SharePoint sites for Project Server and project workspaces are already trusted. Corresponds to the setting on the **Project Server** tab of the **Trust Center** dialog box.|
| _LegacyFileFormats_|Optional|**Integer**|Sets the option for opening or saving files with legacy or non-default file formats. Valid values are 0&ndash;2. Corresponds to the setting on the **Legacy Formats** tab of the **Trust Center** dialog box. Can be one of the constants in the **[PjLegacyFileFormats](Project.PjLegacyFileFormats.md)** enumeration.|

## Return value

 **Boolean**


## Remarks

The **OptionsSecurityEx** method deals with legacy settings for files created in an earlier version of Microsoft Project. To open a specific tab of the **Trust Center** dialog box in Project, use the **[OptionsSecurityTab](Project.Application.OptionsSecurityTab.md)** method.

If an argument is omitted, its default value is specified by the current setting in the **Trust Center** dialog box. Using the **OptionsSecurityEx** method without specifying any arguments displays the **Trust Center** dialog box.

**OptionsSecurityEx** returns **True** if the user clicks **OK** in the **Options** dialog box, or **False** if the user chooses **Cancel**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]