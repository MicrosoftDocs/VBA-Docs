---
title: TextRange.InsertDateTime Method (Publisher)
keywords: vbapb10.chm5308453
f1_keywords:
- vbapb10.chm5308453
ms.prod: publisher
api_name:
- Publisher.TextRange.InsertDateTime
ms.assetid: 1d02471a-f22b-7dad-bcbb-40af3a04d198
ms.date: 06/08/2017
localization_priority: Normal
---


# TextRange.InsertDateTime Method (Publisher)

Returns a  **[TextRange](Publisher.TextRange.md)** object that represents the date and time inserted into a specified text range.


## Syntax

 _expression_. **InsertDateTime**(**_Format_**,  **_InsertAsField_**,  **_InsertAsFullWidth_**,  **_Language_**,  **_Calendar_**)

 _expression_ A variable that represents a  **TextRange** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|Format|Required| **PbDateTimeFormat**|A format for the date and time.|
|InsertAsField|Optional| **Boolean**| **True** for Microsoft Publisher to update date and time whenever opening the publication. Default is **False**.|
|InsertAsFullWidth|Optional| **Boolean**| **True** to insert the specified information as double-byte digits. This argument may not be available to you, depending on the language support (U.S. English, for example) that you have selected or installed. Default is **False**.|
|Language|Optional| **MsoLanguageID**|The language in which to display the date or time.|
|Calendar|Optional| **PbCalendarType**|The calendar type to use when displaying the date or time.|

## Return value

TextRange


## Remarks

The Format parameter can be one of the  **[PbDateTimeFormat](Publisher.PbDateTimeFormat.md)** constants declared in the Microsoft Publisher type library.

The Language parameter can be one of the  ** [MsoLanguageID](Office.MsoLanguageID.md)** constants declared in the Microsoft Office type library.

The Calendar parameter can be one of the  **[PbCalendarType](Publisher.PbCalendarType.md)** constants declared in the Microsoft Publisher type library. The default is **pbCalendarTypeWestern**.


## Example

This example inserts a field for the current date at the cursor position.


```vb
Sub InsertDateField() 
 Selection.TextRange.InsertDateTime Format:=pbDateLong, InsertAsField:=True 
End Sub
```


