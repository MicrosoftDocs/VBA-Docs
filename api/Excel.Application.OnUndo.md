---
title: Application.OnUndo method (Excel)
keywords: vbaxl10.chm133185
f1_keywords:
- vbaxl10.chm133185
ms.prod: excel
api_name:
- Excel.Application.OnUndo
ms.assetid: 12e59bbb-e134-3728-7c8d-629dcda0e908
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.OnUndo method (Excel)

Sets the text of the **Undo** command and the name of the procedure that's run if you choose the **Undo** command after running the procedure that sets this property.


## Syntax

_expression_.**OnUndo** (_Text_, _Procedure_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Required| **String**|The text that appears with the **Undo** command.|
| _Procedure_|Required| **String**|The name of the procedure that's run when you choose the **Undo** command.|

## Remarks

If a procedure doesn't use the **OnUndo** method, the **Undo** command is disabled.

The procedure must use the **[OnRepeat](Excel.Application.OnRepeat.md)** and **OnUndo** methods last to prevent the repeat and undo procedures from being overwritten by subsequent actions in the procedure.


## Example

This example sets the repeat and undo procedures.

```vb
Application.OnRepeat "Repeat VB Procedure", _ 
 "Book1.xls!My_Repeat_Sub" 
Application.OnUndo "Undo VB Procedure", _ 
 "Book1.xls!My_Undo_Sub"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]