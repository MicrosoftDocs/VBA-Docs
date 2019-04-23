---
title: AddFromFile method (VBA Add-In Object Model)
keywords: vbob6.chm1098957
f1_keywords:
- vbob6.chm1098957
ms.prod: office
ms.assetid: 5169e5ee-d5a6-82d3-5a03-dcc84819a752
ms.date: 12/06/2018
localization_priority: Normal
---


# AddFromFile method (VBA Add-In Object Model)

For the **References** collection, adds a reference to a [project](../../Glossary/vbe-glossary.md#project) from a file. For the **[CodeModule](codemodule-object-vba-add-in-object-model.md)** object, adds the contents of a file to a [module](../../Glossary/vbe-glossary.md#module).

## Syntax

_object_.**AddFromFile** (_filename_)

<br/>

The **AddFromFile** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _filename_|Required. A [string expression](../../Glossary/vbe-glossary.md#string-expression) specifying the name of the file that you want to add to the project or module. If the file name isn't found and a path name isn't specified, the directories searched by the **Windows OpenFile** function are searched.|

## Remarks

For the **CodeModule** object, the **AddFromFile** method inserts the contents of the file starting on the line preceding the first [procedure](../../Glossary/vbe-glossary.md#procedure) in the [code module](../../Glossary/vbe-glossary.md#code-module). 

If the module doesn't contain procedures, **AddFromFile** places the contents of the file at the end of the module.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]