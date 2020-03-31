---
title: Execute method (ADO connection)
ROBOTS: INDEX
ms.prod: access
ms.assetid: af190bd9-7167-df59-29ca-a9a86c4957fd
ms.date: 06/08/2019
localization_priority: Normal
---


# Execute method (ADO connection)

**Applies to:** Access 2013 | Access 2016

Executes the specified query, SQL statement, stored procedure, or provider-specific text.

## Syntax

For a non-row-returning command string:

For a row-returning command string:


## Return value

Returns a [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object reference.


## Parameters

-  _CommandText_
    
    - A **String** value that contains the SQL statement, stored procedure, a URL, or provider-specific text to execute. Optionally, table names can be used but only if the provider is SQL aware. For example if a table name of "Customers" is used, ADO will automatically prepend the standard SQL Select syntax to form and pass "SELECT * FROM Customers" as a T-SQL statement to the provider.
    
- _RecordsAffected_
    
    - Optional. A **Long** variable to which the provider returns the number of records that the operation affected.
    
- _Options_
    
    - Optional. A **Long** value that indicates how the provider should evaluate the _CommandText_ argument. Can be a bitmask of one or more [CommandTypeEnum](https://msdn.microsoft.com/library/9ad8f155-88a0-00eb-2855-1e1a2a677437%28Office.15%29.aspx) or [ExecuteOptionEnum](https://msdn.microsoft.com/library/bd6d44a3-e471-7aa0-3e65-6775334de2ff%28Office.15%29.aspx) values.
    

> [!NOTE] 
> Use the **ExecuteOptionEnum** value **adExecuteNoRecords** to improve performance by minimizing internal processing.
> 
> Do not use the **CommandTypeEnum** values of **adCmdFile** or **adCmdTableDirect** with **Execute**. These values can only be used as options with the [Open](https://msdn.microsoft.com/library/87ef19a4-28e1-dec7-ed33-4ae500b9c460%28Office.15%29.aspx) and [Requery](https://msdn.microsoft.com/library/1062d907-979f-020a-b2ed-94e11c0e7d08%28Office.15%29.aspx) methods of a **Recordset**.


## Remarks

Using the **Execute** method on a [Connection](https://msdn.microsoft.com/library/c16023aa-0321-2513-ee71-255d6ffba03d%28Office.15%29.aspx) object executes whatever query you pass to the method in the _CommandText_ argument on the specified connection. If the _CommandText_ argument specifies a row-returning query, any results that the execution generates are stored in a new **Recordset** object. If the command is not intended to return results (for example, an SQL UPDATE query) the provider returns **Nothing** as long as the option **adExecuteNoRecords** is specified; otherwise, Execute returns a closed **Recordset**.

The returned **Recordset** object is always a read-only, forward-only cursor. If you need a **Recordset** object with more functionality, first create a **Recordset** object with the desired property settings, then use the **Recordset** object's [Open](https://msdn.microsoft.com/library/87ef19a4-28e1-dec7-ed33-4ae500b9c460%28Office.15%29.aspx) method to execute the query and return the desired cursor type.

The contents of the  _CommandText_ argument are specific to the provider and can be standard SQL syntax or any special command format that the provider supports.

An [ExecuteComplete](https://msdn.microsoft.com/library/47317d97-e373-32f4-9438-2dff46b8d367%28Office.15%29.aspx) event will be issued when this operation concludes.


> [!NOTE] 
> URLs using the http scheme will automatically invoke the [Microsoft OLE DB Provider for Internet Publishing](https://msdn.microsoft.com/library/5d1e8db5-dabb-0914-e11e-e2eac72bfa77%28Office.15%29.aspx). For more information, see [Absolute and Relative URLs](https://msdn.microsoft.com/library/79a1f793-7154-1c13-7dfe-a1b8cd64e1ea%28Office.15%29.aspx).

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](https://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]