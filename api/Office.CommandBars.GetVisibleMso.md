---
title: CommandBars.GetVisibleMso method (Office)
keywords: vbaof11.chm2020
f1_keywords:
- vbaof11.chm2020
ms.prod: office
api_name:
- Office.CommandBars.GetVisibleMso
ms.assetid: ab916050-e1af-0752-9734-23d0fe27542f
ms.date: 06/08/2017
---


# CommandBars.GetVisibleMso method (Office)

Returns True if the control identified by the  **idMso** parameter is visible.

> [!NOTE] 
>  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `GetVisibleMso`( `_idMso_` )

 _expression_ An expression that returns a [CommandBars](Office.CommandBars.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

## Return value

Boolean


## Example

The following sample returns True if the  **Bold** button is visible.


```vb
Application.CommandBars.GetVisibleMso("Bold")
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/Library-Reference/commandbars-members-office.md)

