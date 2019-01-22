---
title: OfficeDataSourceObject.Move method (Office)
keywords: vbaof11.chm232006
f1_keywords:
- vbaof11.chm232006
ms.prod: office
api_name:
- Office.OfficeDataSourceObject.Move
ms.assetid: cf732e6c-58b3-94a7-5081-3f1350800fd0
ms.date: 01/22/2019
localization_priority: Normal
---


# OfficeDataSourceObject.Move method (Office)

Moves a record in a return set from an **OfficeDataSourceObject** object from one position to another.


## Syntax

_expression_.**Move**(_MsoMoveRow_, _RowNbr_)

_expression_ A variable that represents an **[OfficeDataSourceObject](Office.OfficeDataSourceObject.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MsoMoveRow_|Required|**[MsoMoveRow](office.msomoverow.md)**|A constant specifying which row to move.|
| _RowNbr_|Optional|**Integer**|The number of the destination row.|

## Return value

Integer


## Example

The following example moves the first row in the set of records to the third row.


```vb
oOdso.Move(msoMoveRowFirst, 3)
```


## See also

- [OfficeDataSourceObject object members](overview/library-reference/officedatasourceobject-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]