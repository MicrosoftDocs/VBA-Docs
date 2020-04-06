---
title: ProtectedViewWindows.Open method (PowerPoint)
keywords: vbapp10.chm733004
f1_keywords:
- vbapp10.chm733004
ms.prod: powerpoint
api_name:
- PowerPoint.ProtectedViewWindows.Open
ms.assetid: 864042f4-bfe7-3a70-6428-f7ab166da08d
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindows.Open method (PowerPoint)

Open and return a  **ProtectedViewWindow** object from the **ProtectedViewWindows** collection.


## Syntax

_expression_.**Open** (_FileName_, _ReadPassword_, _OpenAndRepair_)

_expression_ A variable that represents a [ProtectedViewWindows](PowerPoint.ProtectedViewWindows.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to open.|
| _ReadPassword_|Optional|**String**|The password to use for the protected file.|
| _OpenAndRepair_|Optional|**[MSOTRISTATE]**|Indicates whether to repair the file.|

## Return value

 **ProtectedViewWindow** object


## See also


[ProtectedViewWindows Object](PowerPoint.ProtectedViewWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]