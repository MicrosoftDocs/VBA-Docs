---
title: ORDER BY clause (Microsoft Access SQL)
keywords: jetsql40.chm5277571
f1_keywords:
- jetsql40.chm5277571
ms.prod: access
ms.assetid: 9e5e6911-1117-b220-7f11-1ae7f87cbdc0
ms.date: 06/08/2017
localization_priority: Normal
---


# ORDER BY clause (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Sorts a query's resulting records on a specified field or fields in ascending or descending order.

## Syntax

SELECT  _fieldlist_ FROM _table_ WHERE _selectcriteria_ [ORDER BY _field1_ [ASC | DESC ][, _field2_ [ASC | DESC ]][, …]]]

A SELECT statement containing an ORDER BY clause has these parts:



|Part|Description|
|:-----|:-----|
| _fieldlist_|The name of the field or fields to be retrieved along with any field-name aliases, [SQL aggregate functions](https://msdn.microsoft.com/library/8866cd71-0216-25b4-6a6a-02cb7acad9a2%28Office.15%29.aspx), selection predicates ([ALL, DISTINCT, DISTINCTROW, or TOP](all-distinct-distinctrow-top-predicates-microsoft-access-sql.md)), or other [SELECT](https://msdn.microsoft.com/library/a5c9da94-5f9e-0fc0-767a-4117f38a5ef3%28Office.15%29.aspx) statement options.|
| _table_|The name of the table from which records are retrieved. For more information, see the [FROM](from-clause-microsoft-access-sql.md) clause.|
| _selectcriteria_|Selection criteria. If the statement includes a [WHERE](where-clause-microsoft-access-sql.md) clause, the Microsoft Access database engine orders values after applying the WHERE conditions to the records.|
| _field1_, _field2_|The names of the fields on which to sort records.|

## Remarks

ORDER BY is optional. However, if you want your data displayed in sorted order, then you must use ORDER BY.

The default sort order is ascending (A to Z, 0 to 9). Both of the following examples sort employee names in last name order:

```sql
SELECT LastName, FirstName 
FROM Employees 
ORDER BY LastName; 
SELECT LastName, FirstName 
FROM Employees 
ORDER BY LastName ASC;
```

To sort in descending order (Z to A, 9 to 0), add the DESC reserved word to the end of each field you want to sort in descending order. The following example selects salaries and sorts them in descending order:

```sql
SELECT LastName, Salary 
FROM Employees 
ORDER BY Salary DESC, LastName;
```

If you specify a field containing Memo or OLE Object data in the ORDER BY clause, an error occurs. The Microsoft Jet database engine does not sort on fields of these types.

ORDER BY is usually the last item in an SQL statement.

You can include additional fields in the ORDER BY clause. Records are sorted first by the first field listed after ORDER BY. Records that have equal values in that field are then sorted by the value in the second field listed, and so on.
    

## Example

The SQL statement shown in the following example uses the ORDER BY clause to sort records by last name in descending order (Z-A). It calls the EnumFields procedure, which you can find in the SELECT statement example.

```vb
Sub OrderByX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Select the last name and first name values from  
    ' the Employees table, and sort them in descending  
    ' order. 
    Set rst = dbs.OpenRecordset("SELECT LastName, " _ 
        & "FirstName FROM Employees " _ 
        & "ORDER BY LastName DESC;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print recordset contents. 
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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]