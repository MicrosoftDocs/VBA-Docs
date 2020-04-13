---
title: Application.CustomFieldIndicatorDelete method (Project)
keywords: vbapj.chm39
f1_keywords:
- vbapj.chm39
ms.prod: project-server
api_name:
- Project.Application.CustomFieldIndicatorDelete
ms.assetid: 729eafe9-4d1a-07a6-efbc-ab0c94e3af59
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomFieldIndicatorDelete method (Project)

Removes a test condition from a custom field graphical indicator criteria list.


## Syntax

_expression_. `CustomFieldIndicatorDelete`( `_FieldID_`, `_Index_`, `_CriteriaList_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|The custom field. Can be one of the **[PjCustomField](Project.PjCustomField.md)** constants.|
| _Index_|Required|**Integer**|The position of the test condition to delete from the list specified by  **CriteriaList**.|
| _CriteriaList_|Optional|**Long**|The criteria list containing the test condition to be deleted. Can be one of the following  **PjCriteriaList** constants: **pjCriteriaNonSummary**, **pjCriteriaSummary**, or **pjCriteriaProjectSummary**. The default value is **pjCriteriaNonSummary**.|

## Return value

 **Boolean**


## Remarks

The **CustomFieldIndicatorDelete** method returns a trappable error (error code 1004) if the list specified by _CriteriaList_ is read-only because it has been set to inherit values from another list.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]