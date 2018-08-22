---
title: GetSelection Method (VBA Add-In Object Model)
keywords: vbob6.chm1098973
f1_keywords:
- vbob6.chm1098973
ms.prod: office
ms.assetid: f7275ba1-85a3-4939-2ab2-f39e750623f0
ms.date: 06/08/2017
---


# GetSelection Method (VBA Add-In Object Model)



Returns the selection in a [code pane](../../Glossary/vbe-glossary.md#code-pane).

## Syntax

_object_**.GetSelection(**_startline_, _startcol_, _endline_, _endcol_**)**
The  **GetSelection** syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#Long) that returns a value specifying the first line of the selection in the code pane.|
| _startcol_|Required. A  **Long** that returns a value specifying the first column of the selection in the code pane.|
| _endline_|Required. A  **Long** that returns a value specifying the last line of the selection in the code pane.|
| _endcol_|Required. A  **Long** that returns a value specifying the last column of the selection in the code pane.|

## Remarks

When you use the  **GetSelection** method, information is returned in output[arguments](../../Glossary/vbe-glossary.md#argument). As a result, you must pass in [variables](../../Glossary/vbe-glossary.md#variable) because the variables will be modified to contain the information when returned.

