---
title: IBlogExtensibility.GetRecentPosts method (Office)
keywords: vbaof11.chm328004
f1_keywords:
- vbaof11.chm328004
ms.prod: office
api_name:
- Office.IBlogExtensibility.GetRecentPosts
ms.assetid: 460cb59e-c025-8a80-1cdc-99a9c58ec4c0
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogExtensibility.GetRecentPosts method (Office)

Returns the list of the user's last fifteen blog posts that Microsoft Word then displays in the **Open Existing Post** dialog. This method does not actually return the blog post contents.


## Syntax

_expression_.**GetRecentPosts** (_Account_, _ParentWindow_, _Document_, _userName_, _Password_, _PostTitles()_, _PostDates()_, _PostIDs()_)

_expression_ An expression that returns an **[IBlogExtensibility](Office.IBlogExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window that Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents the user's password stored in the registry account settings.|
| _PostTitles()_|Required|**String**|Contains the titles of the last fifteen posts.|
| _PostDates()_|Required|**String**|Contains the dates of the last fifteen posts.|
| _PostIDs()_|Required|**String**|Contains the IDs of the last fifteen posts.|

## See also

- [IBlogExtensibility object members](overview/Library-Reference/iblogextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]