---
title: Avg function (Microsoft Access SQL)
keywords: jetsql40.chm5278823
f1_keywords:
- jetsql40.chm5278823
ms.prod: access
ms.assetid: be955493-a236-2dbe-a08d-2a7f6d113b39
ms.date: 09/21/2018
---


# Avg function (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Calculates the arithmetic mean of a set of values contained in a specified field on a query.

## Syntax

**Avg(** _expr_ **)**

The  _expr_ placeholder represents a string expression identifying the field that contains the numeric data you want to average or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks

The average calculated by **Avg** is the arithmetic mean (the sum of the values divided by the number of values). You could use **Avg**, for example, to calculate average freight cost.

The **Avg** function does not include any **Null** fields in the calculation.

You can use **Avg** in a query expression and in the **SQL** property of a **QueryDef** object or when creating a Recordset object based on an SQL query.


## Example

This example uses the Orders table to calculate the average freight charges for orders with freight charges over $100. 

This example calls the EnumFields procedure, which you can find in the SELECT statement example.

```vb
Sub AvgX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Calculate the average freight charges for orders 
    ' with freight charges over $100.   
    Set rst = dbs.OpenRecordset("SELECT Avg(Freight)" _ 
        & " AS [Average Freight]" _ 
        & " FROM Orders WHERE Freight > 100;") 
    
    ' Populate the Recordset. 
    rst.MoveLast 
    
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 25 
 
    dbs.Close 
 
End Sub
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

