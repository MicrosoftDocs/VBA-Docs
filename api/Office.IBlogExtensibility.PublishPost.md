---
title: IBlogExtensibility.PublishPost method (Office)
keywords: vbaof11.chm328006
f1_keywords:
- vbaof11.chm328006
ms.prod: office
api_name:
- Office.IBlogExtensibility.PublishPost
ms.assetid: 71d2907f-f64e-f318-7f10-095d9a004f03
ms.date: 01/16/2019
localization_priority: Normal
---


# IBlogExtensibility.PublishPost method (Office)

Hands off the current post so that it can be published by the provider.


## Syntax

_expression_.**PublishPost** (_Account_, _ParentWindow_, _Document_, _userName_, _Password_, _xHTML_, _Title_, _DateTime_, _Categories()_, _Draft_, _PostID_, _PublishMessage_)

_expression_ An expression that returns an **[IBlogExtensibility](Office.IBlogExtensibility.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Account_|Required|**String**|Represents the GUID of the account registry key. Blog account settings are stored in the registry at \\HKCU\Software\Microsoft\Office\Common\Blog\Account.|
| _ParentWindow_|Required|**Long**|Contains the HWND for the window that Microsoft Word is calling from.|
| _Document_|Required|**Object**|The current document.|
| _userName_|Required|**String**|Represents the username stored in the registry account settings.|
| _Password_|Required|**String**|Represents the user's password stored in the registry account settings.|
| _xHTML_|Required|**String**|Represents the xHTML of the current document.|
| _Title_|Required|**String**|The title of the post.|
| _DateTime_|Required|**String**|The date the entry was posted.|
| _Categories()_|Required|**String**|A list of categories supported by the provider.|
| _Draft_|Required|**Boolean**|Specifies whether this is a draft version of the post.|
| _PostID_|Required|**String**|The ID of the original post if this post has been republished.|
| _PublishMessage_|Required|**String**|Specifies what is displayed in the publish bar.|

## See also

- [IBlogExtensibility object members](overview/Library-Reference/iblogextensibility-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]