---
title: IDocumentInspector.GetInfo method (Office)
ms.prod: office
api_name:
- Office.IDocumentInspector.GetInfo
ms.assetid: 7242cce4-1b36-107f-ec7c-2512b2e1fba7
ms.date: 06/08/2017
---


# IDocumentInspector.GetInfo method (Office)

Gets information about a custom Document Inspector module.


## Syntax

 _expression_. `GetInfo`( `_Name_`, `_Desc_` )

 _expression_ An expression that returns a [IDocumentInspector](Office.IDocumentInspector.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|Represents the name of the module.|
| _Desc_|Required|**String**|Represents the description of the module.|

## Return value

[HRESULT]

> [!NOTE] 
> The  **IDocumentInspector** object is for the exclusive use of custom Document Inspector module authors and cannot be used with Visual Basic for Applications (VBA).


## See also


[IDocumentInspector Object](Office.IDocumentInspector.md)



[IDocumentInspector Object Members](./overview/Library-Reference/idocumentinspector-members-office.md)

