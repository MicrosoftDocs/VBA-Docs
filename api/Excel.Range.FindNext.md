---
title: Range.FindNext method (Excel)
keywords: vbaxl10.chm144129
f1_keywords:
- vbaxl10.chm144129
ms.prod: excel
api_name:
- Excel.Range.FindNext
ms.assetid: 308c6241-2398-13e6-ba68-17ec713376f6
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.FindNext method (Excel)

Continues a search that was begun with the **[Find](Excel.Range.Find.md)** method. Finds the next cell that matches those same conditions and returns a **Range** object that represents that cell. This does not affect the selection or the active cell.


## Syntax

_expression_.**FindNext** (_After_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _After_|Optional| **Variant**|The cell after which you want to search. This corresponds to the position of the active cell when a search is done from the user interface. Be aware that _After_ must be a single cell in the range.<br/><br/> Remember that the search begins after this cell; the specified cell is not searched until the method wraps back around to this cell. If this argument is not specified, the search starts after the cell in the upper-left corner of the range.|

## Return value

Range


## Remarks

When the search reaches the end of the specified search range, it wraps around to the beginning of the range. To stop a search when this wraparound occurs, save the address of the first found cell, and then test each successive found-cell address against this saved address.


## Example

This example finds all cells in the range A1:A500 that contain the value 2 and changes their value to 5.

```vb
With Worksheets(1).Range("a1:a500")
     Set c = .Find(2, lookin:=xlValues)
     If Not c Is Nothing Then
        firstAddress = c.Address
        Do
            c.Value = 5
            Set c = .FindNext(c)
        If c is Nothing Then
            GoTo DoneFinding
        End If
        Loop While c.Address <> firstAddress
      End If
      DoneFinding:
End With
```

<br/>

This example finds all the cells in the first four columns that contain a constant X, and hides the column that contains the X.

```vb
Sub Hide_Columns()

    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String

    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.Worksheets("Sheet1")
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:D1").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain an X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(What:="X")
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
             
            'Hide the column, and then find the next X.
            Do
                m_rnFind.EntireColumn.Hidden = True
                Set m_rnFind = .FindNext(m_rnFind)
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub
```

<br/>

This example finds all the cells in the first four columns that contain a constant X, and unhides the column that contains the X.

```vb
Sub Unhide_Columns()
    'Excel objects.
    Dim m_wbBook As Workbook
    Dim m_wsSheet As Worksheet
    Dim m_rnCheck As Range
    Dim m_rnFind As Range
    Dim m_stAddress As String
    
    'Initialize the Excel objects.
    Set m_wbBook = ThisWorkbook
    Set m_wsSheet = m_wbBook.Worksheets("Sheet1")
    
    'Search the four columns for any constants.
    Set m_rnCheck = m_wsSheet.Range("A1:D1").SpecialCells(xlCellTypeConstants)
    
    'Retrieve all columns that contain X. If there is at least one, begin the DO/WHILE loop.
    With m_rnCheck
        Set m_rnFind = .Find(What:="X", LookIn:=xlFormulas)
        If Not m_rnFind Is Nothing Then
            m_stAddress = m_rnFind.Address
            
            'Unhide the column, and then find the next X.
            Do
                m_rnFind.EntireColumn.Hidden = False
                Set m_rnFind = .FindNext(m_rnFind)
            Loop While Not m_rnFind Is Nothing And m_rnFind.Address <> m_stAddress
        End If
    End With

End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
