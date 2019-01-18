---
title: IBlogPictureExtensibility.CreatePictureAccount method (Office)
keywords: vbaof11.chm329002
f1_keywords:
- vbaof11.chm329002
ms.prod: office
api_name:
- Office.IBlogPictureExtensibility.CreatePictureAccount
ms.assetid: 8012b234-b8c1-cfc7-7413-b43300fdab76
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogPictureExtensibility.CreatePictureAccount method (Office)

Allows a picture provider to display the user interface needed to guide the user through setting up a picture account.


## Syntax

_expression_.**CreatePictureAccount** (_Account_, _BlogProvider_, _ParentWindow_, _Document_, _userName_, _Password_)

_expression_ An expression that returns an **[IBlogPictureExtensibility](Office.IBlogPictureExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _BlogProvider_|Required|**String**|The ID of the provider.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window that Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents the user's password stored in the registry account settings.|

## See also

- [IBlogPictureExtensibility object members](overview/Library-Reference/iblogpictureextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]