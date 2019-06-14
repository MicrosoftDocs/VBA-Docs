---
title: TextRange.InsertDateTime method (Publisher)
keywords: vbapb10.chm5308453
f1_keywords:
- vbapb10.chm5308453
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertDateTime
ms.assetid: 1d02471a-f22b-7dad-bcbb-40af3a04d198
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.InsertDateTime method (Publisher)

Returns a **TextRange** object that represents the date and time inserted into a specified text range.


## Syntax

_expression_.**InsertDateTime** (_Format_, _InsertAsField_, _InsertAsFullWidth_, _Language_, _Calendar_)

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Format_|Required| **[PbDateTimeFormat](Publisher.PbDateTimeFormat.md)**|A format for the date and time. Can be one of the **PbDateTimeFormat** constants declared in the Microsoft Publisher type library.|
|_InsertAsField_|Optional| **Boolean**| **True** for Publisher to update date and time whenever opening the publication. The default is **False**.|
|_InsertAsFullWidth_|Optional| **Boolean**| **True** to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. The default is **False**.|
|_Language_|Optional| **[MsoLanguageID](Office.MsoLanguageID.md)**|The language in which to display the date or time. Can be one of the **MsoLanguageID** constants declared in the Microsoft Office type library.|
|_Calendar_|Optional| **[PbCalendarType](Publisher.PbCalendarType.md)**|The calendar type to use when displaying the date or time. Can be one of the **PbCalendarType** constants. The default is **pbCalendarTypeWestern**.|


## Return value

TextRange


## Example

This example inserts a field for the current date at the cursor position.

```vb
Sub InsertDateField() 
 Selection.TextRange.InsertDateTime Format:=pbDateLong, InsertAsField:=True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]