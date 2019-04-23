---
title: IDocumentInspector.Inspect method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.Inspect
ms.assetid: 33c767c7-5f28-9cba-6511-513a2efda1a3
ms.date: 01/16/2019
localization_priority: Normal
---


# IDocumentInspector.Inspect method (Office)

Inspects a document for specific information items or document properties by using a custom Document Inspector module.


## Syntax

_expression_.**Inspect** (_Doc_, _Status_, _Result_, _Action_)

_expression_ An expression that returns an **[IDocumentInspector](Office.IDocumentInspector.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required|**Object**|An object representing the container document.|
| _Status_|Required|**[MsoDocInspectorStatus](office.msodocinspectorstatus.md)**|An enumeration that represents the results of the inspection.|
| _Result_|Required|**String**|Contains a list of the information items or document properties found in the document.|
| _Action_|Required|**String**|Indicates to the user what action to take based on the results of the inspection.|

## Return value

[HRESULT]


## Remarks

The **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also

- [IDocumentInspector object members](overview/Library-Reference/idocumentinspector-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]