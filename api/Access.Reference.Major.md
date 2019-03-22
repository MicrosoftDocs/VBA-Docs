---
title: Reference.Major property (Access)
keywords: vbaac10.chm12632
f1_keywords:
- vbaac10.chm12632
ms.prod: access
api_name:
- Access.Reference.Major
ms.assetid: b7aa0cf2-7bdd-58d0-4949-29e3d39be013
ms.date: 03/23/2019
localization_priority: Normal
---


# Reference.Major property (Access)

The **Major** property of a **Reference** object returns a read-only **Long** value indicating the major version number of an application to which you have set a reference.


## Syntax

_expression_.**Major**

_expression_ A variable that represents a **[Reference](Access.Reference.md)** object.


## Remarks

The **Major** property returns the value to the left of the decimal point in a version number. For example, if you've set a reference to an application whose version number is 2.5, the **Major** property returns 2.


## Example

The following example displays a message with information about all the references in the current project.

```vb
Dim r As Reference 
Dim strInfo As String 
 
For Each r In Application.References 
 strInfo = strInfo & r.Name & " " & r.Major & "." & r.Minor & vbCrLf 
Next 
 
 
MsgBox "Current References: " & vbCrLf & strInfo
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]