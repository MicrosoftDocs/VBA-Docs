---
title: IBlogExtensibility.GetUserBlogs method (Office)
keywords: vbaof11.chm328003
f1_keywords:
- vbaof11.chm328003
ms.prod: office
api_name:
- Office.IBlogExtensibility.GetUserBlogs
ms.assetid: 00e76f3d-59f2-8580-6f7e-6df8fe51d345
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogExtensibility.GetUserBlogs method (Office)

Returns the list and details of user blogs associated with the specified account.


## Syntax

_expression_.**GetUserBlogs** (_Account_, _ParentWindow_, _Document_, _userName_, _Password_, _BlogNames()_, _BlogIDs()_, _BlogURLs()_)

_expression_ An expression that returns an **[IBlogExtensibility](Office.IBlogExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window that Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents the user's password stored in the registry account settings.|
| _BlogNames()_|Required|**String**|Contains all blog names under the current account.|
| _BlogIDs()_|Required|**String**|Contains all blog IDs under the current account.|
| _BlogURLs()_|Required|**String**|Contains all blog URLs under the current account.|

## See also

- [IBlogExtensibility object members](overview/Library-Reference/iblogextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]