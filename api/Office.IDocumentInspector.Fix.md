---
title: IDocumentInspector.Fix Method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.Fix
ms.assetid: bf803bd1-5acc-b023-c98b-f21a7f708f6e
ms.date: 06/08/2017
---


# IDocumentInspector.Fix Method (Office)

Performs some action on specific information items or document properties by using a custom Document Inspector module.


## Syntax

 _expression_. `Fix`( `_Doc_`, `_Hwnd_`, `_Status_`, `_Result_` )

 _expression_ An expression that returns a [IDocumentInspector](./Office.IDocumentInspector.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required|**Object**|An object representing the container object.|
| _Hwnd_|Required|**Long**|Unique identifier of the active document window.|
| _Status_|Required|**MsoDocInspectorStatus**|An enumeration that indicates the status of the action.|
| _Result_|Required|**String**|Contains the results of the action.|

## Return value

[HRESULT]


## Remarks

MsoDocInspectorStatus members

> [!NOTE] 
> The  **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also


[IDocumentInspector Object](Office.IDocumentInspector.md)



[IDocumentInspector Object Members](./overview/Library-Reference/idocumentinspector-members-office.md)

