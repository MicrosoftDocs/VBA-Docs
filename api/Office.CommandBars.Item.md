---
title: CommandBars.Item property (Office)
keywords: vbaof11.chm2008
f1_keywords:
- vbaof11.chm2008
ms.prod: office
api_name:
- Office.CommandBars.Item
ms.assetid: bca38d83-67cb-2cba-ddfa-918a5b2ff508
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.Item property (Office)

Gets a **CommandBar** object from the **CommandBars** collection. Read-only.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Item**(_Index_)

_expression_ Required. A variable that represents a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The name or index number of the object to be returned.|

## Example

**Item** is the default member of the object or collection. The following two statements both assign a **CommandBar** object to cmdBar.

```vb
Set cmdBar = CommandBars.Item("Standard") 
Set cmdBar = CommandBars("Standard")
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]