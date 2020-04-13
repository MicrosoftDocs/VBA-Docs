---
title: Variables.Add method (Word)
keywords: vbawd10.chm157614087
f1_keywords:
- vbawd10.chm157614087
ms.prod: word
api_name:
- Word.Variables.Add
ms.assetid: 5c38d785-539b-7e6c-9cd0-cfa48e1aef33
ms.date: 06/08/2017
localization_priority: Normal
---


# Variables.Add method (Word)

Returns a  **Variable** object that represents a variable added to a document.


## Syntax

_expression_.**Add** (_Name_, _Value_)

_expression_ Required. A variable that represents a '[Variables](Word.variables.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the document variable.|
| _Value_|Optional| **Variant**|The value for the document variable.|

## Return value

Variable


## Remarks

Document variables are invisible to the user unless a DOCVARIABLE field is inserted with the appropriate variable name. If you try to add a variable with a name that already exists in the **Variables** collection, an error occurs. To avoid this error, you can enumerate the collection before adding a new variable to it.


## Example

This example adds a variable named Temp to the active document and then inserts a DOCVARIABLE field to display the value in the Temp variable.


```vb
With ActiveDocument 
 .Variables.Add Name:="Temp", Value:="12" 
 .Fields.Add Range:=Selection.Range, _ 
 Type:=wdFieldDocVariable, Text:="Temp" 
End With 
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False
```

This example sets the value of the Blue variable to six. If this variable doesn't already exist, the example adds it to the document and sets it to six.




```vb
For Each aVar In ActiveDocument.Variables 
 If aVar.Name = "Blue" Then num = aVar.Index 
Next aVar 
If num = 0 Then 
 ActiveDocument.Variables.Add Name:="Blue", Value:=6 
Else 
 ActiveDocument.Variables(num).Value = 6 
End If
```

This example stores the user name (from the **Options** dialog box) in the template attached to the active document.




```vb
ScreenUpdating = False 
With ActiveDocument.AttachedTemplate.OpenAsDocument 
 .Variables.Add Name:="UserName", Value:= Application.UserName 
 .Close SaveChanges:=wdSaveChanges 
End With
```


## See also


[Variables Collection Object](Word.variables.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]