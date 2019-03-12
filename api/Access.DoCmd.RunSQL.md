---
title: DoCmd.RunSQL method (Access)
keywords: vbaac10.chm4176
f1_keywords:
- vbaac10.chm4176
ms.prod: access
api_name:
- Access.DoCmd.RunSQL
ms.assetid: 5d61f75a-b220-cc2c-edea-51a6d4f9f106
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.RunSQL method (Access)

The **RunSQL** method carries out the RunSQL action in Visual Basic.


## Syntax

_expression_.**RunSQL** (_SQLStatement_, _UseTransaction_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SQLStatement_|Required|**Variant**|A string expression that's a valid SQL statement for an action query or a data-definition query. It uses an INSERT INTO, DELETE, SELECT...INTO, UPDATE, CREATE TABLE, ALTER TABLE, DROP TABLE, CREATE INDEX, or DROP INDEX statement. Include an IN clause if you want to access another database.|
| _UseTransaction_|Optional|**Variant**|Use **True** (1) to include this query in a transaction. Use **False** (0) if you don't want to use a transaction. If you leave this argument blank, the default (**True**) is assumed.|

## Remarks

You can use the RunSQL action to run a Microsoft Access action query by using the corresponding SQL statement. You can also run a data-definition query.

This method only applies to Access databases.

The maximum length of the _SQLStatement_ argument is 32,768 characters (unlike the _SQLStatement_ action argument in the Macro window, whose maximum length is 256 characters).

## Example

The following example updates the **Employees** table, changing each sales manager's title to Regional Sales Manager.

```vb
Public Sub DoSQL() 
 
    Dim SQL As String 
     
    SQL = "UPDATE Employees" & _ 
          "SET Employees.Title = 'Regional Sales Manager'" & _ 
          "WHERE Employees.Title = 'Sales Manager'" 
 
    DoCmd.RunSQL SQL 
     
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
