---
title: IDocumentInspector.GetInfo method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.GetInfo
ms.assetid: 7242cce4-1b36-107f-ec7c-2512b2e1fba7
ms.date: 01/16/2019
localization_priority: Normal
---


# IDocumentInspector.GetInfo method (Office)

Gets information about a custom Document Inspector module.


## Syntax

_expression_.**GetInfo** (_Name_, _Desc_)

_expression_ An expression that returns an **[IDocumentInspector](Office.IDocumentInspector.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|Represents the name of the module.|
| _Desc_|Required|**String**|Represents the description of the module.|

## Return value

[HRESULT]

## Remarks

The **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also

- [IDocumentInspector object members](overview/Library-Reference/idocumentinspector-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]