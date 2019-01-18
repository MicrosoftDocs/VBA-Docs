---
title: IBlogExtensibility.SetupBlogAccount method (Office)
keywords: vbaof11.chm328002
f1_keywords:
- vbaof11.chm328002
ms.prod: office
api_name:
- Office.IBlogExtensibility.SetupBlogAccount
ms.assetid: 98082a55-3e67-7181-2c7d-2c6979c89ab2
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogExtensibility.SetupBlogAccount method (Office)

Called from the **Choose Account** dialog when the provider's name is chosen in the **Blog Host** drop-down, or when the user requests to change a provider's account in the **Blog Accounts** dialog box.


## Syntax

_expression_.**SetupBlogAccount** (_Account_, _ParentWindow_, _Document_, _NewAccount_, _ShowPictureUI_)

_expression_ An expression that returns an **[IBlogExtensibility](Office.IBlogExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window that Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _NewAccount_|Required|**Boolean**|Indicates whether this is a new account.|
| _ShowPictureUI_|Required|**Boolean**|Indicates whether Word's picture user interface needs to be displayed.|

## See also

- [IBlogExtensibility object members](overview/Library-Reference/iblogextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]