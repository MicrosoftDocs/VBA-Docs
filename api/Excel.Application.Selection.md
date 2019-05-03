---
title: Application.Selection property (Excel)
keywords: vbaxl10.chm183107
f1_keywords:
- vbaxl10.chm183107
ms.prod: excel
api_name:
- Excel.Application.Selection
ms.assetid: f25b5608-035b-983a-545d-d720990c28be
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.Selection property (Excel)

Returns the currently selected object on the active worksheet for an **Application** object. Returns **Nothing** if no objects are selected. Use the **Select** method to set the selection, and use the **[TypeName](../Language/Reference/User-Interface-Help/typename-function.md)** function to discover the kind of object that is selected. 


## Syntax

_expression_.**Selection**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

The returned object type depends on the current selection (for example, if a cell is selected, this property returns a **[Range](Excel.Range(object).md)** object). The **Selection** property returns **Nothing** if nothing is selected.

Using this property with no object qualifier is equivalent to using Application.Selection.


## Example

This example clears the selection on Sheet1 (assuming that the selection is a range of cells).

```vb
Worksheets("Sheet1").Activate 
Selection.Clear
```

<br/>

This example displays the Visual Basic object type of the selection.

```vb
Worksheets("Sheet1").Activate 
MsgBox "The selection object type is " & TypeName(Selection)
```

<br/>

This example displays information about the current selection.

```vb
Sub TestSelection(  )
    Dim str As String
    Select Case TypeName(Selection)
    Case "Nothing"
        str = "No selection made."
    Case "Range"
        str = "You selected the range: " & Selection.Address
    Case "Picture"
        str = "You selected a picture."
    Case Else
        str = "You selected a " & TypeName(Selection) & "."
    End Select
    MsgBox str
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
