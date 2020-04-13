---
title: CodeMask object (Project)
ms.prod: project-server
ms.assetid: 4d0a22f4-fee9-8f4b-a0c0-7bc817ad3f6a
ms.date: 06/08/2017
localization_priority: Normal
---


# CodeMask object (Project)

The **CodeMask** object is a collection of **[CodeMaskLevel](Project.CodeMaskLevel.md)** objects that define the code mask for an outline code in Project.
 


## Example

The following example adds three levels to a code mask.
 

 

```vb
Sub DefineLocationCodeMask(objCodeMask As CodeMask) 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Length:=2, Separator:="." 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Separator:="." 
 
    objCodeMask.Add _ 
        Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 
        Length:=3, Separator:="." 
End Sub
```


## Methods



|Name|
|:-----|
|[Add](Project.CodeMask.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Project.CodeMask.Application.md)|
|[Count](Project.CodeMask.Count.md)|
|[Item](Project.CodeMask.Item.md)|
|[Parent](Project.CodeMask.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]