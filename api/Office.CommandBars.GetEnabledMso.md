---
title: CommandBars.GetEnabledMso Method (Office)
keywords: vbaof11.chm2019
f1_keywords:
- vbaof11.chm2019
ms.prod: office
api_name:
- Office.CommandBars.GetEnabledMso
ms.assetid: 68af6404-53ee-4c69-51fa-4d489736d228
ms.date: 06/08/2017
---


# CommandBars.GetEnabledMso Method (Office)

Returns True if the control identified by the  **idMso** parameter is enabled.


## Syntax

 _expression_. `GetEnabledMso`( `_idMso_` )

 _expression_ An expression that returns a [CommandBars](./Office.CommandBars.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return value

Boolean


## Example

The following sample returns True if the  **Bold** button is enabled.


```vb
Application.CommandBars.GetEnabledMso("Bold")
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/Library-Reference/commandbars-members-office.md)

