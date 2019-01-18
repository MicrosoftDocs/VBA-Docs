---
title: EncryptionProvider.Authenticate method (Office)
keywords: vbaof11.chm327003
f1_keywords:
- vbaof11.chm327003
ms.prod: office
api_name:
- Office.EncryptionProvider.Authenticate
ms.assetid: cb0ecd48-2d37-389c-d041-947b4d9d752a
ms.date: 01/08/2019
localization_priority: Normal
---


# EncryptionProvider.Authenticate method (Office)

Used to determine whether the user has the proper permissions to open the encrypted document.


## Syntax

_expression_.**Authenticate**(_ParentWindow_, _EncryptionData_, _PermissionsMask_)

_expression_ An expression that returns an **[EncryptionProvider](Office.EncryptionProvider.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ParentWindow_|Required|**IUnknown**|Specifies the window that is called to display the encryption settings.|
| _EncryptionData_|Required|**IUnknown**|Contains the encrypted data for the current document.|
| _PermissionsMask_|Required|**Unsigned Integer**|The user interface displayed by the encryption provider add-in.|

## Return value

Long


## Remarks

This is where your COM add-in encryption provider displays whatever user interface is applicable for applying encryption. For example, a password encryption provider would prompt for the user's password.


## See also

- [EncryptionProvider object members](overview/library-reference/encryptionprovider-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]