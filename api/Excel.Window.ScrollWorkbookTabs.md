---
title: Window.ScrollWorkbookTabs method (Excel)
keywords: vbaxl10.chm356107
f1_keywords:
- vbaxl10.chm356107
ms.prod: excel
api_name:
- Excel.Window.ScrollWorkbookTabs
ms.assetid: 5c7c4d74-f125-d67e-2196-14a740afe947
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.ScrollWorkbookTabs method (Excel)

Scrolls through the workbook tabs at the bottom of the window. Doesn't affect the active sheet in the workbook.


## Syntax

_expression_.**ScrollWorkbookTabs** (_Sheets_, _Position_)

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sheets_|Optional| **Variant**|The number of sheets to scroll by. Use a positive number to scroll forward, a negative number to scroll backward, or 0 (zero) to not scroll at all. You must specify _Sheets_ if you don't specify _Position_.|
| _Position_|Optional| **Variant**|Use **xlFirst** to scroll to the first sheet, or use **xlLast** to scroll to the last sheet. You must specify _Position_ if you don't specify _Sheets_.|

## Return value

Variant


## Example

This example scrolls through the workbook tabs to the last sheet in the workbook.

```vb
ActiveWindow.ScrollWorkbookTabs position:=xlLast
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]