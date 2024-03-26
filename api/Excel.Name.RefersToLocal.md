---
title: Name.RefersToLocal property (Excel)
keywords: vbaxl10.chm490085
f1_keywords:
- vbaxl10.chm490085
api_name:
- Excel.Name.RefersToLocal
ms.assetid: e079e8c9-44f9-494e-97aa-2a38c0ec157b
ms.date: 05/01/2019
ms.localizationpriority: medium
---


# Name.RefersToLocal property (Excel)

Returns or sets the formula that the name refers to. The formula is in the language of the user, and it's in A1-style notation, beginning with an equal sign. Read/write **String**.


## Syntax

_expression_.**RefersToLocal**

_expression_ A variable that represents a **[Name](Excel.Name.md)** object.


## Example

This example creates a new worksheet and then inserts a list of all the names in the active workbook, including their formulas (in A1-style notation and in the language of the user).

```vb
Set newSheet = ActiveWorkbook.Worksheets.Add 
i = 1 
For Each nm In ActiveWorkbook.Names 
 newSheet.Cells(i, 1).Value = nm.NameLocal 
 newSheet.Cells(i, 2).Value = "'" & nm.RefersToLocal 
 i = i + 1 
Next
```

## Known issues

This property has a bug. Assigning a localized formula to this property fails. The property expects the same formula syntax as the RefersTo property: US format and list separators. Example:

```vb
Sub Example()
    Dim Nm As Name
    Set Nm = ThisWorkbook.Names("test")
    'If Windows is set to use ; as listseparator, this fails:
    Nm.RefersToLocal = "=SUM(Sheet1!$A$1,Sheet1!$A$3)"
    'If Windows is set to use ; as listseparator, this works:
    Nm.RefersToLocal = "=SUM(Sheet1!$A$1;Sheet1!$A$3)"
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
