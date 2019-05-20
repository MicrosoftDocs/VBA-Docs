---
title: Window.Caption property (Excel)
keywords: vbaxl10.chm356080
f1_keywords:
- vbaxl10.chm356080
ms.prod: excel
api_name:
- Excel.Window.Caption
ms.assetid: d8a5ca13-90b8-d7ce-d041-2cdc544789e5
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.Caption property (Excel)

Returns or sets a **Variant** value that represents the name that appears in the title bar of the document window.


## Syntax

_expression_.**Caption**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

When you set the name, you can use that name as the index to the **[Windows](Excel.Windows.md)** collection (as demonstrated in the example.)


## Example

This example sets the name of the first window in the active workbook to Consolidated Balance Sheet. This name is then used as the index to that window in the **Windows** collection.

```vb
ActiveWorkbook.Windows(1).Caption = "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]