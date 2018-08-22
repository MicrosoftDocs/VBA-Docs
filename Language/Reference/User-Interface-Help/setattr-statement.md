---
title: SetAttr Statement
keywords: vblr6.chm1009017
f1_keywords:
- vblr6.chm1009017
ms.prod: office
ms.assetid: dad85437-6944-a393-9f12-5827b184f42d
ms.date: 06/08/2017
---


# SetAttr Statement

Sets attribute information for a file.

## Syntax

**SetAttr** **_pathname_, _attributes_**

The  **SetAttr** statement syntax has these[named arguments](../../Glossary/vbe-glossary.md#named-argument):


|**Part**|**Description**|
|:-----|:-----|
|**_pathname_**|Required. [String expression](../../Glossary/vbe-glossary.md#String-expression) that specifies a file name — may include directory or folder, and drive.|
|**_attributes_**|Required. [Constant](../../Glossary/vbe-glossary.md#Constant) or[numeric expression](../../Glossary/vbe-glossary.md#numeric-expression), whose sum specifies file attributes.|

 **Settings**
The  **_attributes_**[argument](../../Glossary/vbe-glossary.md#argument) settings are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbNormal**|0|Normal (default).|
|**vbReadOnly**|1|Read-only.|
|**vbHidden**|2|Hidden.|
|**vbSystem**|4|System file. Not available on the Macintosh.|
|**vbArchive**|32|File has changed since last backup.|
|**vbAlias**|64|Specified file name is an alias. Available only on the Macintosh.|

 **Note**  These constants are specified by Visual Basic for Applications. The names can be used anywhere in your code in place of the actual values.

## Remarks

A [run-time error](../../Glossary/vbe-glossary.md#run-time-error) occurs if you try to set the attributes of an open file.

## Example

This example uses the  **SetAttr** statement to set attributes for a file. On the Macintosh, only the constants vbNormal, vbReadOnly, vbHidden and vbAlias are available.


```vb
SetAttr "TESTFILE", vbHidden ' Set hidden attribute. 
SetAttr "TESTFILE", vbHidden + vbReadOnly ' Set hidden and read-only 
 ' attributes. 

```


