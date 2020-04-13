---
title: CodeMaskLevel object (Project)
ms.prod: project-server
api_name:
- Project.CodeMaskLevel
ms.assetid: cef1b15f-c7f1-3b95-49a1-00854a74d9da
ms.date: 06/08/2017
localization_priority: Normal
---


# CodeMaskLevel object (Project)

Represents a level in the code mask of an outline code definition. The **CodeMaskLevel** object is a member of the **[CodeMask](Project.codemask.md)** collection.
 


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
|[Delete](Project.CodeMaskLevel.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.CodeMaskLevel.Application.md)|
|[Index](Project.CodeMaskLevel.Index.md)|
|[Length](Project.CodeMaskLevel.Length.md)|
|[Level](Project.CodeMaskLevel.Level.md)|
|[Parent](Project.CodeMaskLevel.Parent.md)|
|[Separator](Project.CodeMaskLevel.Separator.md)|
|[Sequence](Project.CodeMaskLevel.Sequence.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]