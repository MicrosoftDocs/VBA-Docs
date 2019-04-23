---
title: SharedWorkspaceTask.Save method (Office)
keywords: vbaof11.chm2640011
f1_keywords:
- vbaof11.chm2640011
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Save
ms.assetid: ebddddd5-f42d-5790-7bca-693554982edc
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceTask.Save method (Office)

Uploads changes made programmatically to a shared server.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**Save** (_bstrQueryName_)

_expression_ A variable that represents a **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrQueryName_|Required|**String**|Name of the query used to change the property of the shared workspace link.|

## Remarks

Use the **Save** method to upload changes to the server after changing the properties of a shared workspace task.


## See also

- [SharedWorkspaceTask object members](overview/Library-Reference/sharedworkspacetask-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]