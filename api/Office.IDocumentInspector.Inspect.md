---
title: IDocumentInspector.Inspect Method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.Inspect
ms.assetid: 33c767c7-5f28-9cba-6511-513a2efda1a3
ms.date: 06/08/2017
---


# IDocumentInspector.Inspect Method (Office)

Inspects a document for specific information items or document properties by using a custom Document Inspector module.


## Syntax

 _expression_. `Inspect`( `_Doc_`, `_Status_`, `_Result_`, `_Action_` )

 _expression_ An expression that returns a [IDocumentInspector](./Office.IDocumentInspector.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required|**Object**|An object representing the container document.|
| _Status_|Required|**MsoDocInspectorStatus**|An enumeration that represents the results of the inspection.|
| _Result_|Required|**String**|Contains a list of the information items or document properties found in the document.|
| _Action_|Required|**String**|Indicates to the user what action to take based on the results of the inspection.|

### Return value

[HRESULT]


## Remarks

MsoDocInspectorStatus members

> [!NOTE] 
> The  **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also


[IDocumentInspector Object](Office.IDocumentInspector.md)



[IDocumentInspector Object Members](./overview/Library-Reference/idocumentinspector-members-office.md)

