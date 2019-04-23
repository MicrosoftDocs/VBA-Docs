---
title: Application.RecordMacro method (Excel)
keywords: vbaxl10.chm133195
f1_keywords:
- vbaxl10.chm133195
ms.prod: excel
api_name:
- Excel.Application.RecordMacro
ms.assetid: 8b6c9757-b589-04e6-5650-edfc4104e517
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.RecordMacro method (Excel)

Records code if the macro recorder is on.


## Syntax

_expression_.**RecordMacro** (_BasicCode_, _XlmCode_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BasicCode_|Optional| **Variant**|A string that specifies the Visual Basic code that will be recorded if the macro recorder is recording into a Visual Basic module. The string will be recorded on one line. If the string contains a carriage return (ASCII character 10, or Chr$(10) in code), it will be recorded on more than one line.|
| _XlmCode_|Optional| **Variant**|This argument is ignored.|

## Remarks

The **RecordMacro** method cannot record into the active module (the module in which the **RecordMacro** method exists).

If _BasicCode_ is omitted and the application is recording into Visual Basic, Microsoft Excel will record a suitable Application.Run statement.

To prevent recording (for example, if the user cancels your dialog box), call this function with two empty strings.


## Example

This example records Visual Basic code.

```vb
Application.RecordMacro BasicCode:="Application.Run ""MySub"" "
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]