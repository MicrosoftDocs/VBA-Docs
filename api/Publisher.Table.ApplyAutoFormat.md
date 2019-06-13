---
title: Table.ApplyAutoFormat method (Publisher)
keywords: vbapb10.chm4784137
f1_keywords:
- vbapb10.chm4784137
ms.prod: publisher
api_name:
- Publisher.Table.ApplyAutoFormat
ms.assetid: f792a5f3-0d1c-06de-a030-7a588ca372d2
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.ApplyAutoFormat method (Publisher)

Applies automatic built-in table formatting to a specified table.


## Syntax

_expression_.**ApplyAutoFormat** (_AutoFormat_, _TextFormatting_, _TextAlignment_, _Fill_, _Borders_)

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_AutoFormat_|Required| **[PbTableAutoFormatType](Publisher.PbTableAutoFormatType.md)**|The type of automatic formatting to apply to the specified table. Can be one of the **PbTableAutoFormatType** constants declared in the Microsoft Publisher type library.|
|_TextFormatting_|Optional| **Boolean**| **True** to apply font formatting to the text in the table. Default value is **True**.|
|_TextAlignment_|Optional| **Boolean**| **True** to apply text alignment to the text in the table. Default value is **True**.|
|_Fill_|Optional| **Boolean**| **True** to apply fill formatting to cells in the table. Default value is **True**.|
|_Borders_|Optional| **Boolean**| **True** to apply borders to cells in the table. Default value is **True**.|


## Example

This example applies the checkbook register automatic formatting with fill and borders to the specified table.

```vb
Sub ApplyAutomaticTableFormatting() 
 ActiveDocument.Pages(1).Shapes(1).Table.ApplyAutoFormat _ 
 AutoFormat:=pbTableAutoFormatCheckbookRegister, _ 
 Borders:=False 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]