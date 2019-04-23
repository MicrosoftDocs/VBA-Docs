---
title: Reference.Minor property (Access)
keywords: vbaac10.chm12633
f1_keywords:
- vbaac10.chm12633
ms.prod: access
api_name:
- Access.Reference.Minor
ms.assetid: 7c227db9-9b75-92e5-d32d-e3fda027c145
ms.date: 03/23/2019
localization_priority: Normal
---


# Reference.Minor property (Access)

The **Minor** property of a **Reference** object returns a **Long** value indicating the minor version number of the application to which you have set a reference.


## Syntax

_expression_.**Minor**

_expression_ A variable that represents a **[Reference](Access.Reference.md)** object.


## Remarks

The **Minor** property returns the value to the right of the decimal point in a version number. For example, if you've set a reference to an application whose version number is 2.5, the **Minor** property returns 5.


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