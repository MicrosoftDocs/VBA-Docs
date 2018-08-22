---
title: Lines Property
keywords: vbob6.chm1098980
f1_keywords:
- vbob6.chm1098980
ms.prod: office
api_name:
- Office.Lines
ms.assetid: bd45d817-37c0-c130-7044-4794449505f3
ms.date: 06/08/2017
---


# Lines Property



Returns a string containing the specified number of lines of code.

## Syntax

_object_**.Lines(**_startline_, _count_**) As String**
The  **Lines** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#Long) specifying the line number in which to start.|
| _count_|Required. A  **Long** specifying the number of lines you want to return.|

## Remarks

The [line numbers](../../Glossary/vbe-glossary.md#line-number) in a[code module](../../Glossary/vbe-glossary.md#code-module) begin at 1.

