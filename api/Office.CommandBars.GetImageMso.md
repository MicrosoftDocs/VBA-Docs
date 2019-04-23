---
title: CommandBars.GetImageMso method (Office)
keywords: vbaof11.chm2025
f1_keywords:
- vbaof11.chm2025
ms.prod: office
api_name:
- Office.CommandBars.GetImageMso
ms.assetid: 36261e2b-9cbf-b0b6-5892-63bbb2f93959
ms.date: 01/04/2019
localization_priority: Normal
---


# CommandBars.GetImageMso method (Office)

Returns an **IPictureDisp** object of the control image identified by the **idMso** parameter scaled to the dimensions specified by width and height.

> [!NOTE] 
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**GetImageMso**(_idMso_, _Width_, _Height_)

_expression_ An expression that returns a **[CommandBars](Office.CommandBars.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _idMso_|Required|**String**|Identifier for the control.|
| _Width_|Required|**Integer**|The width of the image.|
| _Height_|Required|**Integer**|The height of the image.|

## Return value

IPictureDisp


## Remarks

The **Width** and **Height** parameters must be between 16 and 128.


## Example

The following sample returns a 32x32 version of the **Paste** icon as an **IPictureDisp** object.


```vb
Application.CommandBars.GetImageMso("Paste", 32, 32)
```


## See also

- [CommandBars object members](overview/library-reference/commandbars-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]