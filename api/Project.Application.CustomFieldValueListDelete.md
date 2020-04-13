---
title: Application.CustomFieldValueListDelete method (Project)
keywords: vbapj.chm42
f1_keywords:
- vbapj.chm42
ms.prod: project-server
api_name:
- Project.Application.CustomFieldValueListDelete
ms.assetid: f8c513b6-2aab-3e42-ca97-7f91f88f5b61
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomFieldValueListDelete method (Project)

Removes an item from the value list for a custom field.


## Syntax

_expression_. `CustomFieldValueListDelete`( `_FieldID_`, `_Index_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the **[PjCustomField](Project.PjCustomField.md)** constants.|
| _Index_|Required|**Integer**|The row number of the value list item to delete from the **Value List** dialog box.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]