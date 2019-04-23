---
title: Names object (Excel)
keywords: vbaxl10.chm487072
f1_keywords:
- vbaxl10.chm487072
ms.prod: excel
api_name:
- Excel.Names
ms.assetid: ffecf89d-7bae-c470-8e37-608857a9de2a
ms.date: 03/30/2019
localization_priority: Normal
---


# Names object (Excel)

A collection of all the **[Name](Excel.Name.md)** objects in the application or workbook.


## Remarks

Each **Name** object represents a defined name for a range of cells. Names can be either built-in names—such as Database, Print_Area, and Auto_Open—or custom names.

The _RefersTo_ argument must be specified in A1-style notation, including dollar signs ($) where appropriate. For example, if cell A10 is selected on Sheet1 and you define a name by using the _RefersTo_ argument "=sheet1!A1:B1", the new name actually refers to cells A10:B10 (because you specified a relative reference). To specify an absolute reference, use "=sheet1!$A$1:$B$1".


## Example

Use the **[Names](Excel.Workbook.Names.md)** property of the **Workbook** object to return the **Names** collection. The following example creates a list of all the names in the active workbook, plus the addresses that they refer to.

```vb
Set nms = ActiveWorkbook.Names 
Set wks = Worksheets(1) 
For r = 1 To nms.Count 
    wks.Cells(r, 2).Value = nms(r).Name 
    wks.Cells(r, 3).Value = nms(r).RefersToRange.Address 
Next
```

<br/>

Use the **Add** method to create a name and add it to the collection. The following example creates a new name that refers to cells A1:C20 on the worksheet named Sheet1.

```vb
Names.Add Name:="test", RefersTo:="=sheet1!$a$1:$c$20"
```

<br/>

Use **Names** (_index_), where _index_ is the name index number or defined name, to return a single **Name** object. The following example deletes the name mySortRange from the active workbook.

```vb
ActiveWorkbook.Names("mySortRange").Delete
```

<br/>

This example uses a named range as the formula for data validation. This example requires the validation data to be on Sheet 2 in the range A2:A100. This validation data is used to validate data entered on Sheet1 in the range D2:D10.

```vb
Sub Add_Data_Validation_From_Other_Worksheet()
'The current Excel workbook and worksheet, a range to define the data to be validated, and the target range
'to place the data in.
Dim wbBook As Workbook
Dim wsTarget As Worksheet
Dim wsSource As Worksheet
Dim rnTarget As Range
Dim rnSource As Range

'Initialize the Excel objects and delete any artifacts from the last time the macro was run.
Set wbBook = ThisWorkbook
With wbBook
    Set wsSource = .Worksheets("Sheet2")
    Set wsTarget = .Worksheets("Sheet1")
    On Error Resume Next
    .Names("Source").Delete
    On Error GoTo 0
End With

'On the source worksheet, create a range in column A of up to 98 cells long, and name it "Source".
With wsSource
    .Range(.Range("A2"), .Range("A100").End(xlUp)).Name = "Source"
End With

'On the target worksheet, create a range 8 cells long in column D.
Set rnTarget = wsTarget.Range("D2:D10")

'Clear out any artifacts from previous macro runs, then set up the target range with the validation data.
With rnTarget
    .ClearContents
    With .Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=Source"
        
'Set up the Error dialog with the appropriate title and message
        .ErrorTitle = "Value Error"
        .ErrorMessage = "You can only choose from the list."
    End With
End With

End Sub
```

## Methods

- [Add](Excel.Names.Add.md)
- [Item](Excel.Names.Item.md)

## Properties

- [Application](Excel.Names.Application.md)
- [Count](Excel.Names.Count.md)
- [Creator](Excel.Names.Creator.md)
- [Parent](Excel.Names.Parent.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]