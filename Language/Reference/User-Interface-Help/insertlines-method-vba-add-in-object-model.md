---
title: InsertLines Method (VBA Add-In Object Model)
keywords: vbob6.chm1098975
f1_keywords:
- vbob6.chm1098975
ms.prod: office
ms.assetid: 6a719fb8-cb52-6a18-c0dc-a8cd09a4814d
ms.date: 06/08/2017
---


# InsertLines Method (VBA Add-In Object Model)



Inserts a line or lines of code at a specified location in a block of code.
<<<<<<< HEAD
 **Syntax**
 _object_**.InsertLines(**_line_, _code_**)**
=======

## Syntax

_object_**.InsertLines(**_line_, _code_**)**
>>>>>>> master
The  **InsertLines** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md) that evaluates to an object in the Applies To list.|
| _line_|Required. A [Long](../../Glossary/vbe-glossary.md) specifying the location at which you want to insert the code.|
| _code_|Required. A [String](../../Glossary/vbe-glossary.md) containing the code you want to insert.|

 **Remarks**
=======
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _line_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the location at which you want to insert the code.|
| _code_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the code you want to insert.|

## Remarks

>>>>>>> master
If the text you insert using the  **InsertLines** method is carriage return-linefeed delimited, it will be inserted as consecutive lines.

