---
title: CommandBars.GetPressedMso Method (Office)
keywords: vbaof11.chm2021
f1_keywords:
- vbaof11.chm2021
ms.prod: office
api_name:
- Office.CommandBars.GetPressedMso
ms.assetid: 97811bb6-cc5c-eccc-9149-76bdfa37541f
ms.date: 06/08/2017
---


# CommandBars.GetPressedMso Method (Office)

Returns a value indicating whether the toggleButton control identified by the  **idMso** parameter is pressed.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `GetPressedMso`( `_idMso_` )

 _expression_ An expression that returns a [CommandBars](./Office.CommandBars.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return value

Boolean


## Example

The following sample returns True when the  **Bold** button is pressed.


```vb
Application.CommandBars.GetPressedMso("Bold") 
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/Library-Reference/commandbars-members-office.md)

