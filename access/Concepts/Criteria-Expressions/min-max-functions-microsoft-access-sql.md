---
title: Min, Max functions (Microsoft Access SQL)
keywords: jetsql40.chm5278826
f1_keywords:
- jetsql40.chm5278826
ms.prod: access
ms.assetid: 5ac77377-1f6a-7b4f-ecbb-5480bc5a3187
ms.date: 09/21/2018
---


# Min, Max functions (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Return the minimum or maximum of a set of values contained in a specified field on a query.

## Syntax

**Min(_expr_)**

**Max(_expr_)**

The  _expr_ placeholder represents a string expression identifying the field that contains the data you want to evaluate or an expression that performs a calculation using the data in that field. Operands in _expr_ can include the name of a table field, a constant, or a function (which can be either intrinsic or user-defined but not one of the other SQL aggregate functions).


## Remarks

You can use **Min** and **Max** to determine the smallest and largest values in a field based on the specified aggregation, or grouping. For example, you could use these functions to return the lowest and highest freight cost. If there is no aggregation specified, the entire table is used.

You can use **Min** and **Max** in a query expression and in the **SQL** property of a **QueryDef** object or when creating a **Recordset** object based on an SQL query.
    

## Example

This example uses the Orders table to return the lowest and highest freight charges for orders shipped to the United Kingdom.

This example calls the EnumFields procedure, which you can find in the SELECT statement example.

```vb
Sub MinMaxX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
     
    ' Return the lowest and highest freight charges for  
    ' orders shipped to the United Kingdom. 
    Set rst = dbs.OpenRecordset("SELECT " _  
        & "Min(Freight) AS [Low Freight], " _ 
        & "Max(Freight)AS [High Freight] " _ 
        & "FROM Orders WHERE ShipCountry = 'UK';") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 12 
 
    dbs.Close 
 
End Sub 

```



### About the contributors

**Link provided by** ![Community Member Icon](../../../images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) the [UtterAccess](https://www.utteraccess.com) community.

- [Record Order](https://www.utteraccess.com/wiki/index.php/Record_Order)

UtterAccess is the premier Microsoft Access wiki and help forum. 

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)