---
title: Execute method (ADO command)
ROBOTS: INDEX
keywords: ado210.chm1231051
f1_keywords:
- ado210.chm1231051
ms.prod: access
ms.assetid: 01812c8c-403e-4428-23f6-86bda747bd0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Execute method (ADO command)

**Applies to:** Access 2013 | Access 2016

Executes the query, SQL statement, or stored procedure specified in the [CommandText](https://msdn.microsoft.com/library/0debec1c-068f-0aea-fce8-e61aa39c5907%28Office.15%29.aspx) property.

## Syntax

For a **Recordset** -returning **Command**:

For a non-recordset-returning **Command**:


## Return value

Returns a [Recordset](https://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) object reference or **Nothing**.


## Parameters

- _RecordsAffected_
    
    - Optional. A **Long** variable to which the provider returns the number of records that the operation affected. The _RecordsAffected_ parameter applies only for action queries or stored procedures. _RecordsAffected_ does not return the number of records returned by a result-returning query or stored procedure. To obtain this information, use the [RecordCount](https://msdn.microsoft.com/library/e3072d10-5bf7-02a8-027e-a9d9a34e3f27%28Office.15%29.aspx) property. The **Execute** method will not return the correct information when used with **adAsyncExecute**, simply because when a command is executed asynchronously, the number of records affected may not yet be known at the time the method returns.
    
- _Parameters_
    
    - Optional. A **Variant** array of parameter values passed with an SQL statement. (Output parameters will not return correct values when passed in this argument.)
    
- _Options_
    
    - Optional. A **Long** value that indicates how the provider should evaluate the [CommandText](https://msdn.microsoft.com/library/0debec1c-068f-0aea-fce8-e61aa39c5907%28Office.15%29.aspx) property of the [Command](https://msdn.microsoft.com/library/64f4ef03-f858-c004-b891-0c96d13a5e6e%28Office.15%29.aspx) object. Can be a bitmask value made using [CommandTypeEnum](https://msdn.microsoft.com/library/9ad8f155-88a0-00eb-2855-1e1a2a677437%28Office.15%29.aspx) and/or [ExecuteOptionEnum](https://msdn.microsoft.com/library/bd6d44a3-e471-7aa0-3e65-6775334de2ff%28Office.15%29.aspx) values. For example, you could use both **adCmdText** and **adExecuteNoRecords** together in combination if you want to have ADO evaluate the value of the **CommandText** property as text and indicate that the command should discard and not return any records that might be generated when the command text executes.
    

## Remarks

Using the **Execute** method on a **Command** object executes the query specified in the **CommandText** property of the object. If the **CommandText** property specifies a row-returning query, any results that the execution generates are stored in a new **Recordset** object. If the command is not a row-returning query, the provider returns a closed **Recordset** object. Some application languages allow you to ignore this return value if no **Recordset** is desired.

If the query has parameters, the current values for the **Command** object's parameters are used unless you override these with parameter values passed with the **Execute** call. You can override a subset of the parameters by omitting new values for some of the parameters when calling the **Execute** method. The order in which you specify the parameters is the same order in which the method passes them. For example, if there were four (or more) parameters and you wanted to pass new values for only the first and fourth parameters, you would pass as the _Parameters_ argument.

> [!NOTE] 
> Output parameters will not return correct values when passed in the  _Parameters_ argument.

An [ExecuteComplete](https://msdn.microsoft.com/library/47317d97-e373-32f4-9438-2dff46b8d367%28Office.15%29.aspx) event will be issued when this operation concludes.

## See also

- [Access for developers forum](https://social.msdn.microsoft.com/Forums/office/home?forum=accessdev)
- [Access help on support.office.com](https://support.office.com/search/results?query=Access)
- [Access help on answers.microsoft.com](https://answers.microsoft.com/)
- [Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)
- [Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]