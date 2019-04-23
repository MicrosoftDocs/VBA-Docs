---
title: For...Next statement (VBA)
keywords: vblr6.chm1008924
f1_keywords:
- vblr6.chm1008924
ms.prod: office
ms.assetid: 53e92bd3-1933-5bc7-f7a4-4e6a3d9bef4a
ms.date: 12/03/2018
localization_priority: Normal
---


# For...Next statement

Repeats a group of [statements](../../Glossary/vbe-glossary.md#statement) a specified number of times.

## Syntax

**For** _counter_ **=** _start_ **To** _end_ [ **Step** _step_ ] <br/>
[ _statements_ ] <br/>
[ **Exit For** ] <br/>
[ _statements_ ] <br/>
**Next** [ _counter_ ]

<br/>

The **For…Next** statement syntax has these parts:

|Part|Description|
|:-----|:-----|
| _counter_|Required. Numeric [variable](../../Glossary/vbe-glossary.md#variable) used as a loop counter. The variable can't be a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) or an [array](../../Glossary/vbe-glossary.md#array) element.|
| _start_|Required. Initial value of _counter_.|
| _end_|Required. Final value of _counter_.|
| _step_|Optional. Amount _counter_ is changed each time through the loop. If not specified, _step_ defaults to one.|
| _statements_|Optional. One or more statements between **For** and **Next** that are executed the specified number of times.|

## Remarks

The _step_ [argument](../../Glossary/vbe-glossary.md#argument) can be either positive or negative. The value of the _step_ argument determines loop processing as follows.

|Value|Loop executes if|
|:-----|:-----|
|Positive or 0| _counter_ <= _end_|
|Negative| _counter_ >= _end_|

After all statements in the loop have executed, _step_ is added to _counter_. At this point, either the statements in the loop execute again (based on the same test that caused the loop to execute initially), or the loop is exited and execution continues with the statement following the **Next** statement.

> [!TIP] 
> Changing the value of _counter_ while inside a loop can make it more difficult to read and debug your code.

Any number of **[Exit For](exit-statement.md)** statements may be placed anywhere in the loop as an alternate way to exit. **Exit For** is often used after evaluating some condition, for example **If...Then**, and transfers control to the statement immediately following **Next**.

You can nest **For...Next** loops by placing one **For...Next** loop within another. Give each loop a unique variable name as its _counter_. The following construction is correct:

```vb
For I = 1 To 10 
 For J = 1 To 10 
 For K = 1 To 10 
 ... 
 Next K 
 Next J 
Next I 

```


> [!NOTE] 
> If you omit _counter_ in a **Next** statement, execution continues as if _counter_ is included. If a **Next** statement is encountered before its corresponding **For** statement, an error occurs.


## Example

This example uses the **For...Next** statement to create a string that contains 10 instances of the numbers 0 through 9, each string separated from the other by a single space. The outer loop uses a loop counter variable that is decremented each time through the loop.


```vb
Dim Words, Chars, MyString 
For Words = 10 To 1 Step -1 ' Set up 10 repetitions. 
 For Chars = 0 To 9 ' Set up 10 repetitions. 
 MyString = MyString & Chars ' Append number to string. 
 Next Chars ' Increment counter 
 MyString = MyString & " " ' Append a space. 
Next Words 

```


## See also

- [Making faster For...Next loops](../../concepts/getting-started/making-faster-fornext-loops.md)
- [Using For...Next statements](../../concepts/getting-started/using-fornext-statements.md)
- [For Each...Next statement](for-eachnext-statement.md)
- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
