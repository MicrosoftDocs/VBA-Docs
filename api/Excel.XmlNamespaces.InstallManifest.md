---
title: XmlNamespaces.InstallManifest method (Excel)
keywords: vbaxl10.chm746078
f1_keywords:
- vbaxl10.chm746078
ms.prod: excel
api_name:
- Excel.XmlNamespaces.InstallManifest
ms.assetid: e462d627-d4d1-b3e9-4d6c-ae7ed91665ad
ms.date: 05/21/2019
localization_priority: Normal
---


# XmlNamespaces.InstallManifest method (Excel)

Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.


## Syntax

_expression_.**InstallManifest** (_Path_, _InstallForAllUsers_)

_expression_ A variable that represents an **[XmlNamespaces](Excel.XmlNamespaces.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path and file name of the XML expansion pack.|
| _InstallForAllUsers_|Optional| **Variant**| **True** installs the XML expansion pack and makes it available to all users on a machine. **False** makes the XML expansion pack available for the current user only. The default is **False**.|

## Remarks

For security purposes, you cannot install an unsigned manifest. For more information about manifests, see the [Smart Document Software Development Kit (SDK)](https://www.microsoft.com/download/details.aspx?id=3929).


## Example

The following example installs the SimpleSample smart document solution on the user's computer and makes it available only to the current user.

> [!NOTE] 
> The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). 

```vb
Application.XMLNamespaces.InstallManifest _ 
 "https://smartdocuments/simplesample/manifest.xml"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]