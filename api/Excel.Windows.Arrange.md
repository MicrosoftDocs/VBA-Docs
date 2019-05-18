---
title: Windows.Arrange method (Excel)
keywords: vbaxl10.chm354073
f1_keywords:
- vbaxl10.chm354073
ms.prod: excel
api_name:
- Excel.Windows.Arrange
ms.assetid: 6b5088ea-6a75-b0df-941f-2032c9cc34a7
ms.date: 05/18/2019
localization_priority: Normal
---


# Windows.Arrange method (Excel)

Arranges the windows on the screen.


## Syntax

_expression_.**Arrange** (_ArrangeStyle_, _ActiveWorkbook_, _SyncHorizontal_, _SyncVertical_)

_expression_ A variable that represents a **[Windows](Excel.Windows.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ArrangeStyle_|Optional| **[XlArrangeStyle](Excel.XlArrangeStyle.md)**|One of the constants of **XlArrangeStyle** specifying how the windows are arranged.|
| _ActiveWorkbook_|Optional| **Variant**| **True** to arrange only the visible windows of the active workbook. **False** to arrange all windows. The default value is **False**.|
| _SyncHorizontal_|Optional| **Variant**|Ignored if _ActiveWorkbook_ is **False** or omitted. **True** to synchronize the windows of the active workbook when scrolling horizontally. **False** to not synchronize the windows. The default value is **False**.|
| _SyncVertical_|Optional| **Variant**|Ignored if _ActiveWorkbook_ is **False** or omitted. **True** to synchronize the windows of the active workbook when scrolling vertically. **False** to not synchronize the windows. The default value is **False**.|

## Return value

Variant

## Example

This example tiles all the windows in the application.

```vb
Application.Windows.Arrange ArrangeStyle:=xlArrangeStyleTiled
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]