---
title: Application.CustomFieldValueListGetItem method (Project)
keywords: vbapj.chm131200
f1_keywords:
- vbapj.chm131200
ms.prod: project-server
api_name:
- Project.Application.CustomFieldValueListGetItem
ms.assetid: 54ab8b15-374a-3c7a-ffe6-bc90b5d4561e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomFieldValueListGetItem method (Project)

Returns the value, description, or phonetic spelling of an item in the value list for a custom field.


## Syntax

_expression_. `CustomFieldValueListGetItem`( `_FieldID_`, `_Item_`, `_Index_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the  **[PjCustomField](Project.PjCustomField.md)** constants.|
| _Item_|Required|**Long**|The information to return. Can be one of the following  **PjValueListItem** constants: **pjValueListValue**, **pjValueListDescription**, or **pjValueListPhonetics**. The default value is **pjValueListValue**.|
| _Index_|Required|**Long**|The row number of the value list item for which to return the information specified with Item.|

## Return value

 **String**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]