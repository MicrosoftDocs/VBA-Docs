---
title: First, Last functions (Microsoft Access SQL)
ROBOTS: INDEX
keywords: jetsql40.chm5278825
f1_keywords:
- jetsql40.chm5278825
ms.prod: access
ms.assetid: 8ea0d390-bb37-003b-fb6c-e15bf2a50718
ms.date: 06/08/2017
localization_priority: Normal
---


# First, Last functions (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Return a field value from the first or last record in the result set returned by a query.

## Syntax

**First**( _expr_ )

**Last**( _expr_ )

The  _expr_ placeholder represents a string expression identifying the field that contains the data you want to use or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks

The **First** and **Last** functions are analogous to the **MoveFirst** and **MoveLast** methods of a DAO Recordset object. They simply return the value of a specified field in the first or last record, respectively, of the result set returned by a query. Because records are usually returned in no particular order (unless the query includes an [ORDER BY](../Structured-Query-Language/order-by-clause-microsoft-access-sql.md) clause), the records returned by these functions will be arbitrary.
    

## Example

This example uses the Employees table to return the values from the LastName field of the first and last records returned from the table.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.

```vb
Sub FirstLastX1() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Return the values from the LastName field of the  
    ' first and last records returned from the table. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        & "First(LastName) as First, " _ 
        & "Last(LastName) as Last FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub 

```

The next example compares using the **First** and **Last** functions with simply using the **Min** and **Max** functions to find the earliest and latest birth dates of Employees.

```vb
Sub FirstLastX2() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Find the earliest and latest birth dates of 
    ' Employees. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        & "First(BirthDate) as FirstBD, " _ 
        & "Last(BirthDate) as LastBD FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
     
    Debug.Print 
 
    ' Find the earliest and latest birth dates of 
    ' Employees. 
    Set rst = dbs.OpenRecordset("SELECT " _ 
        & "Min(BirthDate) as MinBD," _ 
        & "Max(BirthDate) as MaxBD FROM Employees;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub 

```

<a name="AboutContributors"> </a>

## About the contributors

![Community Member Icon](../../../images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The [UtterAccess](https://www.utteraccess.com) community is the premier Microsoft Access wiki and help forum.  

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]