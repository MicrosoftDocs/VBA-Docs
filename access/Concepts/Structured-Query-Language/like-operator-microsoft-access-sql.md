---
title: Like operator (Microsoft Access SQL)
keywords: jetsql40.chm5277589
f1_keywords:
- jetsql40.chm5277589
ms.prod: access
ms.assetid: 70d2ecef-90d7-aff9-398e-8703fb7dfc6e
ms.date: 06/08/2017
---


# Like operator (Microsoft Access SQL)

**Applies to:** Access 2013 | Access 2016

Compares a string expression to a pattern in an SQL expression.

## Syntax

_expression_ **Like** "_pattern_"

The **Like** operator syntax has these parts:

|**Part**|**Description**|
|:-----|:-----|
| _expression_|SQL expression used in a [WHERE clause](where-clause-microsoft-access-sql.md).|
| _pattern_|String or character string literal against which  _expression_ is compared.|

## Remarks

You can use the **Like** operator to find values in a field that match the pattern you specify. For _pattern_, you can specify the complete value (for example, `Like "Smith"`), or you can use wildcard characters to find a range of values (for example, ), or you can use wildcard characters to find a range of values (for example,  `Like "Sm*")`.

In an expression, you can use the **Like** operator to compare a field value to a string expression. For example, if you enter `Like "C*"` in an SQL query, the query returns all field values beginning with the letter C. In a parameter query, you can prompt the user for a pattern to search for.

The following example returns data that begins with the letter P followed by any letter between A and F and three digits:

```vb
Like "P[A-F]###"
```

The following table shows how you can use **Like** to test expressions for different patterns.

|**Kind of match**|**Pattern**|**Match (returns True)**|**No match (returns False)**|
|:-----|:-----|:-----|:-----|
|Multiple characters|a*a|aa, aBa, aBBBa|aBC|
||*ab*|abc, AABB, Xab|aZb, bac|
|Special character|a[*]a|a*a|aaa|
|Multiple characters|ab*|abcdefg, abc|cab, aab|
|Single character|a?a|aaa, a3a, aBa|aBBBa|
|Single digit|a#a|a0a, a1a, a2a|aaa, a10a|
|Range of characters|[a-z]|f, p, j|2, &;|
|Outside a range|[!a-z]|9, &;, %|b, a|
|Not a digit|[!0-9]|A, a, &;, ~|0, 1, 9|
|Combined|a[!b-m]#|An9, az0, a99|abc, aj0|

## Example

This example returns a list of employees whose names begin with the letters A through D. It calls the EnumFields procedure, which you can find in the SELECT statement example.


```vb
Sub LikeX() 
 
    Dim dbs As Database, rst As Recordset 
 
    ' Modify this line to include the path to Northwind 
    ' on your computer. 
    Set dbs = OpenDatabase("Northwind.mdb") 
 
    ' Return a list of employees whose names begin with 
    ' the letters A through D. 
    Set rst = dbs.OpenRecordset("SELECT LastName," _ 
        & " FirstName FROM Employees" _ 
        & " WHERE LastName Like '[A-D]*';") 
 
    ' Populate the Recordset. 
    rst.MoveLast 
 
    ' Call EnumFields to print the contents of the  
    ' Recordset. Pass the Recordset object and desired 
    ' field width. 
    EnumFields rst, 15 
    
    dbs.Close 
 
End Sub
```

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)