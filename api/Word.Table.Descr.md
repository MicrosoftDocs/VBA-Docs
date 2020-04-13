---
title: Table.Descr property (Word)
keywords: vbawd10.chm156303570
f1_keywords:
- vbawd10.chm156303570
ms.prod: word
api_name:
- Word.Table.Descr
ms.assetid: 745b446c-1371-35d5-d6bd-8ad6aa4867fe
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Descr property (Word)

Returns or sets a  **String** that contains a description for the specified table. Read/write.


## Syntax

_expression_. `Descr`

 _expression_ An expression that returns a '[Table](Word.Table.md)' object.


## Remarks

Use the **Descr** property to provide an alternative text description for a table. This property adds text to the **Description** text box on the **Alt Text** tab of the **Table Properties** dialog in Word.

> [!NOTE] 
> Web browsers display alternative text while tables are loading or if they are missing. Web search engines use the alternative text to help find Web pages. Alternative text is also used to assist disabilities.


## Example

The following code example adds an alternative text table description to the first table in the active document.


```vb
Dim doc As Document 
Dim tbl As Table 
 
Set doc = ActiveDocument 
Set tbl = doc.Tables(1) 
 
tbl.Descr = "This is a table description."
```


## See also


[Table Object](Word.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]