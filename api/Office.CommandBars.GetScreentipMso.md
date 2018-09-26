---
title: CommandBars.GetScreentipMso Method (Office)
keywords: vbaof11.chm2023
f1_keywords:
- vbaof11.chm2023
ms.prod: office
api_name:
- Office.CommandBars.GetScreentipMso
ms.assetid: 23411622-2b35-0c0e-9373-9bc75c5e433e
ms.date: 06/08/2017
---


# CommandBars.GetScreentipMso Method (Office)

Returns the screentip of the control identified by the  **idMso** parameter as a String.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. `GetScreentipMso`( `_idMso_` )

 _expression_ An expression that returns a [CommandBars](./Office.CommandBars.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|

### Return Value

String


## Example

The following sample returns the String "Paste".


```vb
Application.CommandBars.GetScreentipMso("Paste")
```


## See also


[CommandBars Object](Office.CommandBars.md)



[CommandBars Object Members](./overview/Library-Reference/commandbars-members-office.md)

