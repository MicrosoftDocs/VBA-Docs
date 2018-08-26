---
title: Resume Statement
keywords: vblr6.chm1009004
f1_keywords:
- vblr6.chm1009004
ms.prod: office
ms.assetid: 57fa9eb3-7e8d-2f7e-20d7-47e468b7836a
<<<<<<< HEAD
ms.date: 06/08/2017
=======
ms.date: 08/24/2018
>>>>>>> master
---


# Resume Statement

Resumes execution after an error-handling routine is finished.

<<<<<<< HEAD
 **Syntax**

 **Resume** [ **0** ]
=======
## Syntax

**Resume** [ **0** ]
>>>>>>> master

 **Resume** **Next**
 **Resume**_line_
The  **Resume** statement syntax can have any of the following forms:


|**Statement**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
|**Resume**|If the error occurred in the same [procedure](../../Glossary/vbe-glossary.md) as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the[statement](../../Glossary/vbe-glossary.md) that last called out of the procedure containing the error-handling routine.|
|**Resume** **Next**|If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or  **On Error Resume Next** statement).|
|**Resume**_line_|Execution resumes at  _line_ specified in the required _line_[argument](../../Glossary/vbe-glossary.md). The  _line_ argument is a[line label](../../Glossary/vbe-glossary.md) or[line number](../../Glossary/vbe-glossary.md) and must be in the same procedure as the error handler.|

 **Remarks**
=======
|**Resume**|If the error occurred in the same [procedure](../../Glossary/vbe-glossary.md#procedure) as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the[statement](../../Glossary/vbe-glossary.md#statement) that last called out of the procedure containing the error-handling routine.|
|**Resume** **Next**|If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or  **On Error Resume Next** statement).|
|**Resume**_line_|Execution resumes at  _line_ specified in the required _line_[argument](../../Glossary/vbe-glossary.md#argument). The  _line_ argument is a[line label](../../Glossary/vbe-glossary.md#line-label) or[line number](../../Glossary/vbe-glossary.md#line-number) and must be in the same procedure as the error handler.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
>>>>>>> master
If you use a  **Resume** statement anywhere except in an error-handling routine, an error occurs.

## Example

This example uses the  **Resume** statement to end error handling in a procedure, and then resume execution with the statement that caused the error. Error number 55 is generated to illustrate using the **Resume** statement.


```vb
Sub ResumeStatementDemo() 
 On Error GoTo ErrorHandler ' Enable error-handling routine. 
 Open "TESTFILE" For Output As #1 ' Open file for output. 
 Kill "TESTFILE" ' Attempt to delete open file. 
 Exit Sub ' Exit Sub to avoid error handler. 
ErrorHandler: ' Error-handling routine. 
 Select Case Err.Number ' Evaluate error number. 
<<<<<<< HEAD
 Case 55 ' "File already open" error. 
 Close #1 ' Close open file. 
 Case Else 
 ' Handle other situations here.... 
 End Select 
 Resume ' Resume execution at same line 
 ' that caused the error. 
=======
  Case 55 ' "File already open" error. 
   Close #1 ' Close open file. 
  Case Else 
   ' Handle other situations here.... 
 End Select 
 Resume ' Resume execution at same line that caused the error. 
>>>>>>> master
End Sub
```


