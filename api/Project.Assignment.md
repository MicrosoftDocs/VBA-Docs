---
title: Assignment object (Project)
ms.prod: project-server
api_name:
- Project.Assignment
ms.assetid: bfb9a505-7818-0a86-9d4b-f19a0ff465d3
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment object (Project)

Represents an assignment for a task or resource. The **Assignment** object is a member of an **[Assignments](Project.assignments.md)** or an **[OverAllocatedAssignments](./Project.overallocatedassignments.md)** collection.


## Example

 **Using the Assignment Object**

Use  **Assignments** (_index_), where _index_ is the assignment index number, to return a single **Assignment** object. The following example displays the name of the first resource assigned to the specified task.




```vb
MsgBox ActiveProject.Tasks(1).Assignments(1).ResourceName
```

 **Using the Assignments Collection**

Use the **[Assignments](./Project.Task.Assignments.md)** property to return an **Assignments** collection. The following example displays all the resources assigned to the specified task.




```vb
Dim A As Assignment 
 
For Each A In ActiveProject.Tasks(1).Assignments 
 MsgBox A.ResourceName 
Next A
```

Use the **[Add](./Project.Assignments.Add.md)** method to add an **Assignment** object to the **Assignments** collection. The following example adds a resource identified by the number 212 as a new assignment for the specified task.




```vb
ActiveProject.Tasks(1).Assignments.Add ResourceID:=212
```


## Methods



|Name|
|:-----|
|[AppendNotes](./Project.Assignment.AppendNotes.md)|
|[Delete](./Project.Assignment.Delete.md)|
|[EnterpriseTeamMember](./Project.Assignment.EnterpriseTeamMember.md)|
|[Replan](./Project.Assignment.Replan.md)|
|[TimeScaleData](./Project.Assignment.TimeScaleData.md)|

## Properties



|Name|
|:-----|
|[ActualCost](./Project.Assignment.ActualCost.md)|
|[ActualFinish](./Project.Assignment.ActualFinish.md)|
|[ActualOvertimeCost](./Project.Assignment.ActualOvertimeCost.md)|
|[ActualOvertimeWork](./Project.Assignment.ActualOvertimeWork.md)|
|[ActualStart](./Project.Assignment.ActualStart.md)|
|[ActualWork](./Project.Assignment.ActualWork.md)|
|[ACWP](./Project.Assignment.ACWP.md)|
|[Application](./Project.Assignment.Application.md)|
|[Baseline10BudgetCost](./Project.Assignment.Baseline10BudgetCost.md)|
|[Baseline10BudgetWork](./Project.Assignment.Baseline10BudgetWork.md)|
|[Baseline10Cost](./Project.Assignment.Baseline10Cost.md)|
|[Baseline10Finish](./Project.Assignment.Baseline10Finish.md)|
|[Baseline10Start](./Project.Assignment.Baseline10Start.md)|
|[Baseline10Work](./Project.Assignment.Baseline10Work.md)|
|[Baseline1BudgetCost](./Project.Assignment.Baseline1BudgetCost.md)|
|[Baseline1BudgetWork](./Project.Assignment.Baseline1BudgetWork.md)|
|[Baseline1Cost](./Project.Assignment.Baseline1Cost.md)|
|[Baseline1Finish](./Project.Assignment.Baseline1Finish.md)|
|[Baseline1Start](./Project.Assignment.Baseline1Start.md)|
|[Baseline1Work](./Project.Assignment.Baseline1Work.md)|
|[Baseline2BudgetCost](./Project.Assignment.Baseline2BudgetCost.md)|
|[Baseline2BudgetWork](./Project.Assignment.Baseline2BudgetWork.md)|
|[Baseline2Cost](./Project.Assignment.Baseline2Cost.md)|
|[Baseline2Finish](./Project.Assignment.Baseline2Finish.md)|
|[Baseline2Start](./Project.Assignment.Baseline2Start.md)|
|[Baseline2Work](./Project.Assignment.Baseline2Work.md)|
|[Baseline3BudgetCost](./Project.Assignment.Baseline3BudgetCost.md)|
|[Baseline3BudgetWork](./Project.Assignment.Baseline3BudgetWork.md)|
|[Baseline3Cost](./Project.Assignment.Baseline3Cost.md)|
|[Baseline3Finish](./Project.Assignment.Baseline3Finish.md)|
|[Baseline3Start](./Project.Assignment.Baseline3Start.md)|
|[Baseline3Work](./Project.Assignment.Baseline3Work.md)|
|[Baseline4BudgetCost](./Project.Assignment.Baseline4BudgetCost.md)|
|[Baseline4BudgetWork](./Project.Assignment.Baseline4BudgetWork.md)|
|[Baseline4Cost](./Project.Assignment.Baseline4Cost.md)|
|[Baseline4Finish](./Project.Assignment.Baseline4Finish.md)|
|[Baseline4Start](./Project.Assignment.Baseline4Start.md)|
|[Baseline4Work](./Project.Assignment.Baseline4Work.md)|
|[Baseline5BudgetCost](./Project.Assignment.Baseline5BudgetCost.md)|
|[Baseline5BudgetWork](./Project.Assignment.Baseline5BudgetWork.md)|
|[Baseline5Cost](./Project.Assignment.Baseline5Cost.md)|
|[Baseline5Finish](./Project.Assignment.Baseline5Finish.md)|
|[Baseline5Start](./Project.Assignment.Baseline5Start.md)|
|[Baseline5Work](./Project.Assignment.Baseline5Work.md)|
|[Baseline6BudgetCost](./Project.Assignment.Baseline6BudgetCost.md)|
|[Baseline6BudgetWork](./Project.Assignment.Baseline6BudgetWork.md)|
|[Baseline6Cost](./Project.Assignment.Baseline6Cost.md)|
|[Baseline6Finish](./Project.Assignment.Baseline6Finish.md)|
|[Baseline6Start](./Project.Assignment.Baseline6Start.md)|
|[Baseline6Work](./Project.Assignment.Baseline6Work.md)|
|[Baseline7BudgetCost](./Project.Assignment.Baseline7BudgetCost.md)|
|[Baseline7BudgetWork](./Project.Assignment.Baseline7BudgetWork.md)|
|[Baseline7Cost](./Project.Assignment.Baseline7Cost.md)|
|[Baseline7Finish](./Project.Assignment.Baseline7Finish.md)|
|[Baseline7Start](./Project.Assignment.Baseline7Start.md)|
|[Baseline7Work](./Project.Assignment.Baseline7Work.md)|
|[Baseline8BudgetCost](./Project.Assignment.Baseline8BudgetCost.md)|
|[Baseline8BudgetWork](./Project.Assignment.Baseline8BudgetWork.md)|
|[Baseline8Cost](./Project.Assignment.Baseline8Cost.md)|
|[Baseline8Finish](./Project.Assignment.Baseline8Finish.md)|
|[Baseline8Start](./Project.Assignment.Baseline8Start.md)|
|[Baseline8Work](./Project.Assignment.Baseline8Work.md)|
|[Baseline9BudgetCost](./Project.Assignment.Baseline9BudgetCost.md)|
|[Baseline9BudgetWork](./Project.Assignment.Baseline9BudgetWork.md)|
|[Baseline9Cost](./Project.Assignment.Baseline9Cost.md)|
|[Baseline9Finish](./Project.Assignment.Baseline9Finish.md)|
|[Baseline9Start](./Project.Assignment.Baseline9Start.md)|
|[Baseline9Work](./Project.Assignment.Baseline9Work.md)|
|[BaselineBudgetCost](./Project.Assignment.BaselineBudgetCost.md)|
|[BaselineBudgetWork](./Project.Assignment.BaselineBudgetWork.md)|
|[BaselineCost](./Project.Assignment.BaselineCost.md)|
|[BaselineFinish](./Project.Assignment.BaselineFinish.md)|
|[BaselineStart](./Project.Assignment.BaselineStart.md)|
|[BaselineWork](./Project.Assignment.BaselineWork.md)|
|[BCWP](./Project.Assignment.BCWP.md)|
|[BCWS](./Project.Assignment.BCWS.md)|
|[BookingType](./Project.Assignment.BookingType.md)|
|[BudgetCost](./Project.Assignment.BudgetCost.md)|
|[BudgetWork](./Project.Assignment.BudgetWork.md)|
|[Confirmed](./Project.Assignment.Confirmed.md)|
|[Cost](./Project.Assignment.Cost.md)|
|[Cost1](./Project.Assignment.Cost1.md)|
|[Cost10](./Project.Assignment.Cost10.md)|
|[Cost2](./Project.Assignment.Cost2.md)|
|[Cost3](./Project.Assignment.Cost3.md)|
|[Cost4](./Project.Assignment.Cost4.md)|
|[Cost5](./Project.Assignment.Cost5.md)|
|[Cost6](./Project.Assignment.Cost6.md)|
|[Cost7](./Project.Assignment.Cost7.md)|
|[Cost8](./Project.Assignment.Cost8.md)|
|[Cost9](./Project.Assignment.Cost9.md)|
|[CostRateTable](./Project.Assignment.CostRateTable.md)|
|[CostVariance](./Project.Assignment.CostVariance.md)|
|[Created](./Project.Assignment.Created.md)|
|[CV](./Project.Assignment.CV.md)|
|[Date1](./Project.Assignment.Date1.md)|
|[Date10](./Project.Assignment.Date10.md)|
|[Date2](./Project.Assignment.Date2.md)|
|[Date3](./Project.Assignment.Date3.md)|
|[Date4](./Project.Assignment.Date4.md)|
|[Date5](./Project.Assignment.Date5.md)|
|[Date6](./Project.Assignment.Date6.md)|
|[Date7](./Project.Assignment.Date7.md)|
|[Date8](./Project.Assignment.Date8.md)|
|[Date9](./Project.Assignment.Date9.md)|
|[Delay](./Project.Assignment.Delay.md)|
|[Duration1](./Project.Assignment.Duration1.md)|
|[Duration10](./Project.Assignment.Duration10.md)|
|[Duration2](./Project.Assignment.Duration2.md)|
|[Duration3](./Project.Assignment.Duration3.md)|
|[Duration4](./Project.Assignment.Duration4.md)|
|[Duration5](./Project.Assignment.Duration5.md)|
|[Duration6](./Project.Assignment.Duration6.md)|
|[Duration7](./Project.Assignment.Duration7.md)|
|[Duration8](./Project.Assignment.Duration8.md)|
|[Duration9](./Project.Assignment.Duration9.md)|
|[Finish](./Project.Assignment.Finish.md)|
|[Finish1](./Project.Assignment.Finish1.md)|
|[Finish10](./Project.Assignment.Finish10.md)|
|[Finish2](./Project.Assignment.Finish2.md)|
|[Finish3](./Project.Assignment.Finish3.md)|
|[Finish4](./Project.Assignment.Finish4.md)|
|[Finish5](./Project.Assignment.Finish5.md)|
|[Finish6](./Project.Assignment.Finish6.md)|
|[Finish7](./Project.Assignment.Finish7.md)|
|[Finish8](./Project.Assignment.Finish8.md)|
|[Finish9](./Project.Assignment.Finish9.md)|
|[FinishVariance](./Project.Assignment.FinishVariance.md)|
|[FixedMaterialAssignment](./Project.Assignment.FixedMaterialAssignment.md)|
|[Flag1](./Project.Assignment.Flag1.md)|
|[Flag10](./Project.Assignment.Flag10.md)|
|[Flag11](./Project.Assignment.Flag11.md)|
|[Flag12](./Project.Assignment.Flag12.md)|
|[Flag13](./Project.Assignment.Flag13.md)|
|[Flag14](./Project.Assignment.Flag14.md)|
|[Flag15](./Project.Assignment.Flag15.md)|
|[Flag16](./Project.Assignment.Flag16.md)|
|[Flag17](./Project.Assignment.Flag17.md)|
|[Flag18](./Project.Assignment.Flag18.md)|
|[Flag19](./Project.Assignment.Flag19.md)|
|[Flag2](./Project.Assignment.Flag2.md)|
|[Flag20](./Project.Assignment.Flag20.md)|
|[Flag3](./Project.Assignment.Flag3.md)|
|[Flag4](./Project.Assignment.Flag4.md)|
|[Flag5](./Project.Assignment.Flag5.md)|
|[Flag6](./Project.Assignment.Flag6.md)|
|[Flag7](./Project.Assignment.Flag7.md)|
|[Flag8](./Project.Assignment.Flag8.md)|
|[Flag9](./Project.Assignment.Flag9.md)|
|[Guid](./Project.Assignment.Guid.md)|
|[Hyperlink](./Project.Assignment.Hyperlink.md)|
|[HyperlinkAddress](./Project.Assignment.HyperlinkAddress.md)|
|[HyperlinkHREF](./Project.Assignment.HyperlinkHREF.md)|
|[HyperlinkScreenTip](./Project.Assignment.HyperlinkScreenTip.md)|
|[HyperlinkSubAddress](./Project.Assignment.HyperlinkSubAddress.md)|
|[Index](./Project.Assignment.Index.md)|
|[LevelingDelay](./Project.Assignment.LevelingDelay.md)|
|[LinkedFields](./Project.Assignment.LinkedFields.md)|
|[Notes](./Project.Assignment.Notes.md)|
|[Number1](./Project.Assignment.Number1.md)|
|[Number10](./Project.Assignment.Number10.md)|
|[Number11](./Project.Assignment.Number11.md)|
|[Number12](./Project.Assignment.Number12.md)|
|[Number13](./Project.Assignment.Number13.md)|
|[Number14](./Project.Assignment.Number14.md)|
|[Number15](./Project.Assignment.Number15.md)|
|[Number16](./Project.Assignment.Number16.md)|
|[Number17](./Project.Assignment.Number17.md)|
|[Number18](./Project.Assignment.Number18.md)|
|[Number19](./Project.Assignment.Number19.md)|
|[Number2](./Project.Assignment.Number2.md)|
|[Number20](./Project.Assignment.Number20.md)|
|[Number3](./Project.Assignment.Number3.md)|
|[Number4](./Project.Assignment.Number4.md)|
|[Number5](./Project.Assignment.Number5.md)|
|[Number6](./Project.Assignment.Number6.md)|
|[Number7](./Project.Assignment.Number7.md)|
|[Number8](./Project.Assignment.Number8.md)|
|[Number9](./Project.Assignment.Number9.md)|
|[Overallocated](./Project.Assignment.Overallocated.md)|
|[OvertimeCost](./Project.Assignment.OvertimeCost.md)|
|[OvertimeWork](./Project.Assignment.OvertimeWork.md)|
|[Owner](./Project.Assignment.Owner.md)|
|[Parent](./Project.Assignment.Parent.md)|
|[Peak](./Project.Assignment.Peak.md)|
|[PercentWorkComplete](./Project.Assignment.PercentWorkComplete.md)|
|[Project](./Project.Assignment.Project.md)|
|[RegularWork](./Project.Assignment.RegularWork.md)|
|[RemainingCost](./Project.Assignment.RemainingCost.md)|
|[RemainingOvertimeCost](./Project.Assignment.RemainingOvertimeCost.md)|
|[RemainingOvertimeWork](./Project.Assignment.RemainingOvertimeWork.md)|
|[RemainingWork](./Project.Assignment.RemainingWork.md)|
|[Resource](./Project.Assignment.Resource.md)|
|[ResourceGuid](./Project.Assignment.ResourceGuid.md)|
|[ResourceID](./Project.Assignment.ResourceID.md)|
|[ResourceName](./Project.Assignment.ResourceName.md)|
|[ResourceRequestType](./Project.Assignment.ResourceRequestType.md)|
|[ResourceType](./Project.Assignment.ResourceType.md)|
|[ResourceUniqueID](./Project.Assignment.ResourceUniqueID.md)|
|[ResponsePending](./Project.Assignment.ResponsePending.md)|
|[Start](./Project.Assignment.Start.md)|
|[Start1](./Project.Assignment.Start1.md)|
|[Start10](./Project.Assignment.Start10.md)|
|[Start2](./Project.Assignment.Start2.md)|
|[Start3](./Project.Assignment.Start3.md)|
|[Start4](./Project.Assignment.Start4.md)|
|[Start5](./Project.Assignment.Start5.md)|
|[Start6](./Project.Assignment.Start6.md)|
|[Start7](./Project.Assignment.Start7.md)|
|[Start8](./Project.Assignment.Start8.md)|
|[Start9](./Project.Assignment.Start9.md)|
|[StartVariance](./Project.Assignment.StartVariance.md)|
|[Summary](./Project.Assignment.Summary.md)|
|[SV](./Project.Assignment.SV.md)|
|[Task](./Project.Assignment.Task.md)|
|[TaskGuid](./Project.Assignment.TaskGuid.md)|
|[TaskID](./Project.Assignment.TaskID.md)|
|[TaskName](./Project.Assignment.TaskName.md)|
|[TaskOutlineNumber](./Project.Assignment.TaskOutlineNumber.md)|
|[TaskSummaryName](./Project.Assignment.TaskSummaryName.md)|
|[TaskUniqueID](./Project.Assignment.TaskUniqueID.md)|
|[TeamStatusPending](./Project.Assignment.TeamStatusPending.md)|
|[Text1](./Project.Assignment.Text1.md)|
|[Text10](./Project.Assignment.Text10.md)|
|[Text11](./Project.Assignment.Text11.md)|
|[Text12](./Project.Assignment.Text12.md)|
|[Text13](./Project.Assignment.Text13.md)|
|[Text14](./Project.Assignment.Text14.md)|
|[Text15](./Project.Assignment.Text15.md)|
|[Text16](./Project.Assignment.Text16.md)|
|[Text17](./Project.Assignment.Text17.md)|
|[Text18](./Project.Assignment.Text18.md)|
|[Text19](./Project.Assignment.Text19.md)|
|[Text2](./Project.Assignment.Text2.md)|
|[Text20](./Project.Assignment.Text20.md)|
|[Text21](./Project.Assignment.Text21.md)|
|[Text22](./Project.Assignment.Text22.md)|
|[Text23](./Project.Assignment.Text23.md)|
|[Text24](./Project.Assignment.Text24.md)|
|[Text25](./Project.Assignment.Text25.md)|
|[Text26](./Project.Assignment.Text26.md)|
|[Text27](./Project.Assignment.Text27.md)|
|[Text28](./Project.Assignment.Text28.md)|
|[Text29](./Project.Assignment.Text29.md)|
|[Text3](./Project.Assignment.Text3.md)|
|[Text30](./Project.Assignment.Text30.md)|
|[Text4](./Project.Assignment.Text4.md)|
|[Text5](./Project.Assignment.Text5.md)|
|[Text6](./Project.Assignment.Text6.md)|
|[Text7](./Project.Assignment.Text7.md)|
|[Text8](./Project.Assignment.Text8.md)|
|[Text9](./Project.Assignment.Text9.md)|
|[UniqueID](./Project.Assignment.UniqueID.md)|
|[Units](./Project.Assignment.Units.md)|
|[UpdateNeeded](./Project.Assignment.UpdateNeeded.md)|
|[VAC](./Project.Assignment.VAC.md)|
|[WBS](./Project.Assignment.WBS.md)|
|[Work](./Project.Assignment.Work.md)|
|[WorkContour](./Project.Assignment.WorkContour.md)|
|[WorkVariance](./Project.Assignment.WorkVariance.md)|
|[Compliant](./Project.assignment.compliant.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]