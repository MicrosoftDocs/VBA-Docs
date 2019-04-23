---
title: HAVING clause (Microsoft Access SQL)
keywords: jetsql40.chm5277570
f1_keywords:
- jetsql40.chm5277570
ms.prod: access
ms.assetid: 4fc4655b-c8a6-2ca2-509e-ac98d9a1c776
ms.date: 06/08/2017
localization_priority: Normal
---


# HAVING clause (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Specifies which grouped records are displayed in a [SELECT](https://msdn.microsoft.com/library/a5c9da94-5f9e-0fc0-767a-4117f38a5ef3%28Office.15%29.aspx) statement with a GROUP BY clause. After [GROUP BY](group-by-clause-microsoft-access-sql.md) combines records, HAVING displays any records grouped by the GROUP BY clause that satisfy the conditions of the HAVING clause.

## Syntax

SELECT  _fieldlist_ FROM _table_ WHERE _selectcriteria_ GROUP BY _groupfieldlist_ [HAVING _groupcriteria_ ]

A SELECT statement containing a HAVING clause has these parts:

|Part|Description|
|:-----|:-----|
| _fieldlist_|The name of the field or fields to be retrieved along with any field-name aliases, [SQL aggregate functions](https://msdn.microsoft.com/library/8866cd71-0216-25b4-6a6a-02cb7acad9a2%28Office.15%29.aspx), selection predicates ([ALL, DISTINCT, DISTINCTROW, or TOP](all-distinct-distinctrow-top-predicates-microsoft-access-sql.md)), or other SELECT statement options.|
| _table_|The name of the table from which records are retrieved. For more information, see the [FROM](from-clause-microsoft-access-sql.md) clause.|
| _selectcriteria_|Selection criteria. If the statement includes a [WHERE](where-clause-microsoft-access-sql.md) clause, the Microsoft Access database engine groups values after applying the WHERE conditions to the records.|
| _groupfieldlist_|The names of up to 10 fields used to group records. The order of the field names in  _groupfieldlist_ determines the grouping levels from the highest to the lowest level of grouping.|
| _groupcriteria_|An expression that determines which grouped records to display.|

## Remarks

HAVING is optional.

HAVING is similar to WHERE, which determines which records are selected. After records are grouped with GROUP BY, HAVING determines which records are displayed:

```sql
SELECT CategoryID, 
Sum(UnitsInStock) 
FROM Products 
GROUP BY CategoryID 
HAVING Sum(UnitsInStock) > 100 And Like "BOS*";
```

A HAVING clause can contain up to 40 expressions linked by logical operators, such as **And** and **Or**.


## Example

This example selects the job titles assigned to more than one employee in the Washington region. It calls the EnumFields procedure, which you can find in the SELECT statement example.

```vb
Sub HavingX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Select the job titles assigned to more than one  
    ' employee in the Washington region.  
    Set rst = dbs.OpenRecordset("SELECT Title, " _ 
        & "Count(Title) as Total FROM Employees " _ 
        & "WHERE Region = 'WA' " _ 
        & "GROUP BY Title HAVING Count(Title) > 1;") 
     
    ' Populate the Recordset. 
    rst.MoveLast 
     
    ' Call EnumFields to print recordset contents. 
    EnumFields rst, 25 
 
    dbs.Close 
 
End Sub 

```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]