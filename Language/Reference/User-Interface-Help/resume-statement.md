---
title: Resume statement (VBA)
keywords: vblr6.chm1009004
f1_keywords:
- vblr6.chm1009004
ms.prod: office
ms.assetid: 57fa9eb3-7e8d-2f7e-20d7-47e468b7836a
ms.date: 12/03/2018
localization_priority: Normal
---


# Resume statement

Resumes execution after an error-handling routine is finished.

## Syntax

**Resume** [ **0** ]<br/>
**Resume Next**<br/>
**Resume** _line_

<br/>

The **Resume** statement syntax can have any of the following forms:

|Statement|Description|
|:-----|:-----|
|**Resume**|If the error occurred in the same [procedure](../../Glossary/vbe-glossary.md#procedure) as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the [statement](../../Glossary/vbe-glossary.md#statement) that last called out of the procedure containing the error-handling routine.|
|**Resume Next**|If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or the **[On Error Resume Next](on-error-statement.md)** statement).|
|**Resume** _line_|Execution resumes at the _line_ specified in the required _line_ [argument](../../Glossary/vbe-glossary.md#argument). The _line_ argument is a [line label](../../Glossary/vbe-glossary.md#line-label) or [line number](../../Glossary/vbe-glossary.md#line-number) and must be in the same procedure as the error handler.|


## Remarks

If you use a **Resume** statement anywhere except in an error-handling routine, an error occurs.

## Example

This example uses the **Resume** statement to end error handling in a procedure, and then resume execution with the statement that caused the error. Error number 55 is generated to illustrate using the **Resume** statement.


```vb
Sub ResumeStatementDemo() 
 On Error GoTo ErrorHandler ' Enable error-handling routine. 
 Open "TESTFILE" For Output As #1 ' Open file for output. 
 Kill "TESTFILE" ' Attempt to delete open file. 
 Exit Sub ' Exit Sub to avoid error handler. 
ErrorHandler: ' Error-handling routine. 
 Select Case Err.Number ' Evaluate error number. 
  Case 55 ' "File already open" error. 
   Close #1 ' Close open file. 
  Case Else 
   ' Handle other situations here.... 
 End Select 
 Resume ' Resume execution at same line that caused the error. 
End Sub
```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
