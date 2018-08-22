---
title: Print Method
keywords: vblr6.chm1010081
f1_keywords:
- vblr6.chm1010081
ms.prod: office
api_name:
- Office.Print
ms.assetid: 489447fa-e0ea-404a-10f2-23dcd9a8e41a
ms.date: 06/08/2017
---


# Print Method



Prints text in the  **Immediate** window.

## Syntax

_object_**.Print** [ _outputlist_ ]
The  **Print** method syntax has the following object qualifier and part:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Optional. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _outputlist_|Optional. [Expression](../../Glossary/vbe-glossary.md#Expression) or list of expressions to print. If omitted, a blank line is printed.|

The  _outputlist_[argument](../../Glossary/vbe-glossary.md#argument) has the following syntax and parts:
{ **Spc(**_n_**)** |**Tab(**_n_**)** } _expression charpos_


|**Part**|**Description**|
|:-----|:-----|
|**Spc(**_n_**)**|Optional. Used to insert space characters in the output, where  _n_ is the number of space characters to insert.|
|**Tab(**_n_**)**|Optional. Used to position the insertion point at an absolute column number where  _n_ is the column number. Use **Tab** with no argument to position the insertion point at the beginning of the next[print zone](../../Glossary/vbe-glossary.md#print-zone).|
| _expression_|Optional. [Numeric expression](../../Glossary/vbe-glossary.md#Numeric-expression) or[string expression](../../Glossary/vbe-glossary.md#string-expression) to print.|
| _charpos_|Optional. Specifies the insertion point for the next character. Use a semicolon (**;**) to position the insertion point immediately following the last character displayed. Use **Tab(**_n_**)** to position the insertion point at an absolute column number. Use **Tab** with no argument to position the insertion point at the beginning of the next print zone. If _charpos_ is omitted, the next character is printed on the next line.|

## Remarks

Multiple expressions can be separated with either a space or a semicolon.
All data printed to the  **Immediate** window is properly formatted using the decimal separator for the[locale](../../Glossary/vbe-glossary.md#locale) settings specified for your system. The[keywords](../../Glossary/vbe-glossary.md#keyword) are output in the appropriate language for the[host application](../../Glossary/vbe-glossary.md#host-application).
For [Boolean](../../Glossary/vbe-glossary.md#Boolean) data, either `True` or `False` is printed. The **True** and **False** keywords are translated according to the locale setting for the host application.
[Date](../../Glossary/vbe-glossary.md#Date) data is written using the standard short date format recognized by your system. When either the date or the time component is missing or zero, only the data provided is written.
Nothing is written if  _outputlist_ data is[Empty](../../Glossary/vbe-glossary.md#Empty). However, if  _outputlist_ data is[Null](../../Glossary/vbe-glossary.md#Null),  `Null` is output. The **Null** keyword is appropriately translated when it is output.
For error data, the output is written as  `Error errorcode`. The  **Error** keyword is appropriately translated when it is output.
The  _object_ is required if the method is used outside a[module](../../Glossary/vbe-glossary.md#module) having a default display space. For example an error occurs if the method is called in a[standard module](../../Glossary/vbe-glossary.md#standard-module) without specifying an _object_, but if called in a form module, _outputlist_ is displayed on the form.

 **Note**  Because the  **Print** method typically prints with proportionally-spaced characters, there is no correlation between the number of characters printed and the number of fixed-width columns those characters occupy. For example, a wide letter, such as a "W", occupies more than one fixed-width column, and a narrow letter, such as an "i", occupies less. To allow for cases where wider than average characters are used, your tabular columns must be positioned far enough apart. Alternatively, you can print using a fixed-pitch font (such as Courier) to ensure that each character uses only one column.


## Example

Using the  **Print** method, this example displays the value of the variable `MyVar` in the **Immediate** window. Note that the **Print** method only applies to objects that can display text.


```vb
Dim MyVar
MyVar = "Come see me in the Immediate pane."
Debug.Print MyVar

```


