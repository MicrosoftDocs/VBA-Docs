---
title: AddFromFile Method (VBA Add-In Object Model)
keywords: vbob6.chm1098957
f1_keywords:
- vbob6.chm1098957
ms.prod: office
ms.assetid: 5169e5ee-d5a6-82d3-5a03-dcc84819a752
ms.date: 06/08/2017
---


# AddFromFile Method (VBA Add-In Object Model)



For the  **References** collection, adds a reference to a [project](../../Glossary/vbe-glossary.md#project) from a file. For the **CodeModule** object, adds the contents of a file to a [module](../../Glossary/vbe-glossary.md#module).

## Syntax

_object_**.AddFromFile(**_filename_**)**
The  **AddFromFile** syntax has these parts:


|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _filename_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression) specifying the name of the file you want to add to the project or module. If the file name isn't found and a path name isn't specified, the directories searched by the **Windows OpenFile** function are searched.|

## Remarks

For the  **CodeModule** object, the **AddFromFile** method inserts the contents of the file starting on the line preceding the first[procedure](../../Glossary/vbe-glossary.md#procedure) in the [code module](../../Glossary/vbe-glossary.md#code-module). If the module doesn't contain procedures,  **AddFromFile** places the contents of the file at the end of the module.

