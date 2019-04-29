---
title: Range.Copy method (Excel)
keywords: vbaxl10.chm144104
f1_keywords:
- vbaxl10.chm144104
ms.prod: excel
api_name:
- Excel.Range.Copy
ms.assetid: ac5207ac-6be5-3c7e-2c61-67954a59e9df
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Copy method (Excel)

Copies the range to the specified range or to the Clipboard.


## Syntax

_expression_.**Copy** (_Destination_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Destination_|Optional| **Variant**|Specifies the new range to which the specified range will be copied. If this argument is omitted, Microsoft Excel copies the range to the Clipboard.|

## Return value

Variant


## Example

The following code example copies the formulas in cells A1:D4 on Sheet1 into cells E5:H8 on Sheet2.


```vb
Worksheets("Sheet1").Range("A1:D4").Copy _ 
    destination:=Worksheets("Sheet2").Range("E5")
```



 **Sample code provided by:** Bill Jelen, [MrExcel.com](https://www.mrexcel.com/)

The following code example inspects the value in column D for each row on Sheet1. If the value in column D equals "A" the entire row is copied onto SheetA, in the next empty row. If the value equals "B" the row is copied onto SheetB.




```vb
Public Sub CopyRows() 
    Sheets("Sheet1").Select 
    ' Find the last row of data 
    FinalRow = Cells(Rows.Count, 1).End(xlUp).Row 
    ' Loop through each row 
    For x = 2 To FinalRow 
        ' Decide if to copy based on column D 
        ThisValue = Cells(x, 4).Value 
        If ThisValue = "A" Then 
            Cells(x, 1).Resize(1, 33).Copy 
            Sheets("SheetA").Select 
            NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1 
            Cells(NextRow, 1).Select 
            ActiveSheet.Paste 
            Sheets("Sheet1").Select 
        ElseIf ThisValue = "B" Then 
            Cells(x, 1).Resize(1, 33).Copy 
            Sheets("SheetB").Select 
            NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1 
            Cells(NextRow, 1).Select 
            ActiveSheet.Paste 
            Sheets("Sheet1").Select 
        End If 
    Next x 
End Sub
```


### About the contributor

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## See also


[Range Object](Excel.Range(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
