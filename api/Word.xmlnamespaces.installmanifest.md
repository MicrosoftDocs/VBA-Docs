---
title: XMLNamespaces.InstallManifest method (Word)
keywords: vbawd10.chm248971366
f1_keywords:
- vbawd10.chm248971366
ms.prod: word
ms.assetid: ab8805f3-5009-7322-5bd7-3005af630c5d
ms.date: 06/08/2017
localization_priority: Normal
---


# XMLNamespaces.InstallManifest method (Word)

Installs the specified XML expansion pack on the user's computer, making an XML smart document solution available to one or more users.


## Syntax

_expression_.**InstallManifest** (_Path_, _InstallForAllUsers_)

 _expression_ An expression that represents an 'XMLNamespaces' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The path and file name of the XML expansion pack.|
| _InstallForAllUsers_|Optional| **Boolean**| **True** installs the XML expansion pack and makes it available to all users on a computer. **False** makes the XML expansion pack available for the current user only. Default is **False**.|

## Remarks

For security purposes, you cannot install an unsigned manifest. For more information about manifests, see the Smart Document Software Development Kit (SDK) on the Microsoft Developer Network (MSDN) Web site.


## Example

The following code example installs the SimpleSample smart document solution on the user's computer and makes it available only to the current user.


> [!NOTE] 
> The SimpleSample schema is included in the Smart Document Software Development Kit (SDK). For more information, refer to the Smart Document SDK.


```vb
Application.XMLNamespaces.InstallManifest _ 
 "https://smartdocuments/simplesample/manifest.xml"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]