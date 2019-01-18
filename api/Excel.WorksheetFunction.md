---
title: WorksheetFunction object (Excel)
keywords: vbaxl10.chm136072
f1_keywords:
- vbaxl10.chm136072
ms.prod: excel
api_name:
- Excel.WorksheetFunction
ms.assetid: 7b1d5639-363d-632c-2cf0-2232562646b6
ms.date: 08/28/2018
localization_priority: Priority
---


# WorksheetFunction object (Excel)

Used as a container for Microsoft Excel worksheet functions that can be called from Visual Basic.


## Example

Use the **[WorksheetFunction](Excel.Application.WorksheetFunction.md)** property to return the **WorksheetFunction** object. The following example displays the result of applying the **Min** worksheet function to the range A1:C10.


```vb
Set myRange = Worksheets("Sheet1").Range("A1:C10") 
answer = Application.WorksheetFunction.Min(myRange) 
MsgBox answer
```

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](https://www.mrexcel.com/store/index.php?l=product_detail&amp;p=1)

This example uses the  **CountA** worksheet function to determine how many cells in column A contain a value. For this example, the values in column A should be text. This example does a spell check on each value in column A, and if the value is spelled incorrectly, inserts the text "Wrong" into column B; otherwise, it inserts the text "OK" into column B.




```vb
Sub StartSpelling()
   'Set up your variables
   Dim iRow As Integer
   
   'And define your error handling routine.
   On Error GoTo ERRORHANDLER
   
   'Go through all the cells in column A, and perform a spellcheck on the value.
   'If the value is spelled incorrectly, write "Wrong" in column B, otherwise write "OK".
   For iRow = 1 To WorksheetFunction.CountA(Columns(1))
      If Application.CheckSpelling( _
         Cells(iRow, 1).Value, , True) = False Then
         Cells(iRow, 2).Value = "Wrong"
      Else
         Cells(iRow, 2).Value = "OK"
      End If
   Next iRow
   Exit Sub

    'Error handling routine.
ERRORHANDLER:
    MsgBox "The spell check feature is not installed!"
    
End Sub
```


### About the contributor

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


## Methods



|Name|
|:-----|
|[AccrInt](Excel.WorksheetFunction.AccrInt.md)|
|[AccrIntM](Excel.WorksheetFunction.AccrIntM.md)|
|[Acos](Excel.WorksheetFunction.Acos.md)|
|[Acosh](Excel.WorksheetFunction.Acosh.md)|
|[Acot](Excel.worksheetfunction.acot.md)|
|[Acoth](Excel.worksheetfunction.acoth.md)|
|[Aggregate](Excel.WorksheetFunction.Aggregate.md)|
|[AmorDegrc](Excel.WorksheetFunction.AmorDegrc.md)|
|[AmorLinc](Excel.WorksheetFunction.AmorLinc.md)|
|[And](Excel.WorksheetFunction.And.md)|
|[Arabic](Excel.worksheetfunction.arabic.md)|
|[Asc](Excel.WorksheetFunction.Asc.md)|
|[Asin](Excel.WorksheetFunction.Asin.md)|
|[Asinh](Excel.WorksheetFunction.Asinh.md)|
|[Atan2](Excel.WorksheetFunction.Atan2.md)|
|[Atanh](Excel.WorksheetFunction.Atanh.md)|
|[AveDev](Excel.WorksheetFunction.AveDev.md)|
|[Average](Excel.WorksheetFunction.Average.md)|
|[AverageIf](Excel.WorksheetFunction.AverageIf.md)|
|[AverageIfs](Excel.WorksheetFunction.AverageIfs.md)|
|[BahtText](Excel.WorksheetFunction.BahtText.md)|
|[Base](Excel.worksheetfunction.base.md)|
|[BesselI](Excel.WorksheetFunction.BesselI.md)|
|[BesselJ](Excel.WorksheetFunction.BesselJ.md)|
|[BesselK](Excel.WorksheetFunction.BesselK.md)|
|[BesselY](Excel.WorksheetFunction.BesselY.md)|
|[Beta_Dist](Excel.WorksheetFunction.Beta_Dist.md)|
|[Beta_Inv](Excel.WorksheetFunction.Beta_Inv.md)|
|[BetaDist](Excel.WorksheetFunction.BetaDist.md)|
|[BetaInv](Excel.WorksheetFunction.BetaInv.md)|
|[Bin2Dec](Excel.WorksheetFunction.Bin2Dec.md)|
|[Bin2Hex](Excel.WorksheetFunction.Bin2Hex.md)|
|[Bin2Oct](Excel.WorksheetFunction.Bin2Oct.md)|
|[Binom_Dist](Excel.WorksheetFunction.Binom_Dist.md)|
|[Binom_Dist_Range](Excel.worksheetfunction.binom_dist_range.md)|
|[Binom_Inv](Excel.WorksheetFunction.Binom_Inv.md)|
|[BinomDist](Excel.WorksheetFunction.BinomDist.md)|
|[Bitand](Excel.worksheetfunction.bitand.md)|
|[Bitlshift](Excel.worksheetfunction.bitlshift.md)|
|[Bitor](Excel.worksheetfunction.bitor.md)|
|[Bitrshift](Excel.worksheetfunction.bitrshift.md)|
|[Bitxor](Excel.worksheetfunction.bitxor.md)|
|[Ceiling](Excel.WorksheetFunction.Ceiling.md)|
|[Ceiling_Math](Excel.worksheetfunction.ceiling_math.md)|
|[Ceiling_Precise](Excel.WorksheetFunction.Ceiling_Precise.md)|
|[ChiDist](Excel.WorksheetFunction.ChiDist.md)|
|[ChiInv](Excel.WorksheetFunction.ChiInv.md)|
|[ChiSq_Dist](Excel.WorksheetFunction.ChiSq_Dist.md)|
|[ChiSq_Dist_RT](Excel.WorksheetFunction.ChiSq_Dist_RT.md)|
|[ChiSq_Inv](Excel.WorksheetFunction.ChiSq_Inv.md)|
|[ChiSq_Inv_RT](Excel.WorksheetFunction.ChiSq_Inv_RT.md)|
|[ChiSq_Test](Excel.WorksheetFunction.ChiSq_Test.md)|
|[ChiTest](Excel.WorksheetFunction.ChiTest.md)|
|[Choose](Excel.WorksheetFunction.Choose.md)|
|[Clean](Excel.WorksheetFunction.Clean.md)|
|[Combin](Excel.WorksheetFunction.Combin.md)|
|[Combina](Excel.worksheetfunction.combina.md)|
|[Complex](Excel.WorksheetFunction.Complex.md)|
|[Confidence](Excel.WorksheetFunction.Confidence.md)|
|[Confidence_Norm](Excel.WorksheetFunction.Confidence_Norm.md)|
|[Confidence_T](Excel.WorksheetFunction.Confidence_T.md)|
|[Convert](Excel.WorksheetFunction.Convert.md)|
|[Correl](Excel.WorksheetFunction.Correl.md)|
|[Cosh](Excel.WorksheetFunction.Cosh.md)|
|[Cot](Excel.worksheetfunction.cot.md)|
|[Coth](Excel.worksheetfunction.coth.md)|
|[Count](Excel.WorksheetFunction.Count.md)|
|[CountA](Excel.WorksheetFunction.CountA.md)|
|[CountBlank](Excel.WorksheetFunction.CountBlank.md)|
|[CountIf](Excel.WorksheetFunction.CountIf.md)|
|[CountIfs](Excel.WorksheetFunction.CountIfs.md)|
|[CoupDayBs](Excel.WorksheetFunction.CoupDayBs.md)|
|[CoupDays](Excel.WorksheetFunction.CoupDays.md)|
|[CoupDaysNc](Excel.WorksheetFunction.CoupDaysNc.md)|
|[CoupNcd](Excel.WorksheetFunction.CoupNcd.md)|
|[CoupNum](Excel.WorksheetFunction.CoupNum.md)|
|[CoupPcd](Excel.WorksheetFunction.CoupPcd.md)|
|[Covar](Excel.WorksheetFunction.Covar.md)|
|[Covariance_P](Excel.WorksheetFunction.Covariance_P.md)|
|[Covariance_S](Excel.WorksheetFunction.Covariance_S.md)|
|[CritBinom](Excel.WorksheetFunction.CritBinom.md)|
|[Csc](Excel.worksheetfunction.csc.md)|
|[Csch](Excel.worksheetfunction.csch.md)|
|[CumIPmt](Excel.WorksheetFunction.CumIPmt.md)|
|[CumPrinc](Excel.WorksheetFunction.CumPrinc.md)|
|[DAverage](Excel.WorksheetFunction.DAverage.md)|
|[Days](Excel.worksheetfunction.days.md)|
|[Days360](Excel.WorksheetFunction.Days360.md)|
|[Db](Excel.WorksheetFunction.Db.md)|
|[Dbcs](Excel.WorksheetFunction.Dbcs.md)|
|[DCount](Excel.WorksheetFunction.DCount.md)|
|[DCountA](Excel.WorksheetFunction.DCountA.md)|
|[Ddb](Excel.WorksheetFunction.Ddb.md)|
|[Dec2Bin](Excel.WorksheetFunction.Dec2Bin.md)|
|[Dec2Hex](Excel.WorksheetFunction.Dec2Hex.md)|
|[Dec2Oct](Excel.WorksheetFunction.Dec2Oct.md)|
|[Decimal](Excel.worksheetfunction.decimal.md)|
|[Degrees](Excel.WorksheetFunction.Degrees.md)|
|[Delta](Excel.WorksheetFunction.Delta.md)|
|[DevSq](Excel.WorksheetFunction.DevSq.md)|
|[DGet](Excel.WorksheetFunction.DGet.md)|
|[Disc](Excel.WorksheetFunction.Disc.md)|
|[DMax](Excel.WorksheetFunction.DMax.md)|
|[DMin](Excel.WorksheetFunction.DMin.md)|
|[Dollar](Excel.WorksheetFunction.Dollar.md)|
|[DollarDe](Excel.WorksheetFunction.DollarDe.md)|
|[DollarFr](Excel.WorksheetFunction.DollarFr.md)|
|[DProduct](Excel.WorksheetFunction.DProduct.md)|
|[DStDev](Excel.WorksheetFunction.DStDev.md)|
|[DStDevP](Excel.WorksheetFunction.DStDevP.md)|
|[DSum](Excel.WorksheetFunction.DSum.md)|
|[Duration](Excel.WorksheetFunction.Duration.md)|
|[DVar](Excel.WorksheetFunction.DVar.md)|
|[DVarP](Excel.WorksheetFunction.DVarP.md)|
|[EDate](Excel.WorksheetFunction.EDate.md)|
|[Effect](Excel.WorksheetFunction.Effect.md)|
|[EncodeURL](Excel.worksheetfunction.encodeurl.md)|
|[EoMonth](Excel.WorksheetFunction.EoMonth.md)|
|[Erf](Excel.WorksheetFunction.Erf.md)|
|[Erf_Precise](Excel.WorksheetFunction.Erf_Precise.md)|
|[ErfC](Excel.WorksheetFunction.ErfC.md)|
|[ErfC_Precise](Excel.WorksheetFunction.ErfC_Precise.md)|
|[Even](Excel.WorksheetFunction.Even.md)|
|[Expon_Dist](Excel.WorksheetFunction.Expon_Dist.md)|
|[ExponDist](Excel.WorksheetFunction.ExponDist.md)|
|[F_Dist](Excel.WorksheetFunction.F_Dist.md)|
|[F_Dist_RT](Excel.WorksheetFunction.F_Dist_RT.md)|
|[F_Inv](Excel.WorksheetFunction.F_Inv.md)|
|[F_Inv_RT](Excel.WorksheetFunction.F_Inv_RT.md)|
|[F_Test](Excel.WorksheetFunction.F_Test.md)|
|[Fact](Excel.WorksheetFunction.Fact.md)|
|[FactDouble](Excel.WorksheetFunction.FactDouble.md)|
|[FDist](Excel.WorksheetFunction.FDist.md)|
|[FilterXML](Excel.worksheetfunction.filterxml.md)|
|[Find](Excel.WorksheetFunction.Find.md)|
|[FindB](Excel.WorksheetFunction.FindB.md)|
|[FInv](Excel.WorksheetFunction.FInv.md)|
|[Fisher](Excel.WorksheetFunction.Fisher.md)|
|[FisherInv](Excel.WorksheetFunction.FisherInv.md)|
|[Fixed](Excel.WorksheetFunction.Fixed.md)|
|[Floor](Excel.WorksheetFunction.Floor.md)|
|[Floor_Math](Excel.worksheetfunction.floor_math.md)|
|[Floor_Precise](Excel.WorksheetFunction.Floor_Precise.md)|
|[Forecast](Excel.WorksheetFunction.Forecast.md)|
|[Frequency](Excel.WorksheetFunction.Frequency.md)|
|[FTest](Excel.WorksheetFunction.FTest.md)|
|[Fv](Excel.WorksheetFunction.Fv.md)|
|[FVSchedule](Excel.WorksheetFunction.FVSchedule.md)|
|[Gamma](Excel.worksheetfunction.gamma.md)|
|[Gamma_Dist](Excel.WorksheetFunction.Gamma_Dist.md)|
|[Gamma_Inv](Excel.WorksheetFunction.Gamma_Inv.md)|
|[GammaDist](Excel.WorksheetFunction.GammaDist.md)|
|[GammaInv](Excel.WorksheetFunction.GammaInv.md)|
|[GammaLn](Excel.WorksheetFunction.GammaLn.md)|
|[GammaLn_Precise](Excel.WorksheetFunction.GammaLn_Precise.md)|
|[Gauss](Excel.worksheetfunction.gauss.md)|
|[Gcd](Excel.WorksheetFunction.Gcd.md)|
|[GeoMean](Excel.WorksheetFunction.GeoMean.md)|
|[GeStep](Excel.WorksheetFunction.GeStep.md)|
|[Growth](Excel.WorksheetFunction.Growth.md)|
|[HarMean](Excel.WorksheetFunction.HarMean.md)|
|[Hex2Bin](Excel.WorksheetFunction.Hex2Bin.md)|
|[Hex2Dec](Excel.WorksheetFunction.Hex2Dec.md)|
|[Hex2Oct](Excel.WorksheetFunction.Hex2Oct.md)|
|[HLookup](Excel.WorksheetFunction.HLookup.md)|
|[HypGeom_Dist](Excel.WorksheetFunction.HypGeom_Dist.md)|
|[HypGeomDist](Excel.WorksheetFunction.HypGeomDist.md)|
|[IfError](Excel.WorksheetFunction.IfError.md)|
|[IfNa](Excel.worksheetfunction.ifna.md)|
|[ImAbs](Excel.WorksheetFunction.ImAbs.md)|
|[Imaginary](Excel.WorksheetFunction.Imaginary.md)|
|[ImArgument](Excel.WorksheetFunction.ImArgument.md)|
|[ImConjugate](Excel.WorksheetFunction.ImConjugate.md)|
|[ImCos](Excel.WorksheetFunction.ImCos.md)|
|[ImCosh](Excel.worksheetfunction.imcosh.md)|
|[ImCot](Excel.worksheetfunction.imcot.md)|
|[ImCsc](Excel.worksheetfunction.imcsc.md)|
|[ImCsch](Excel.worksheetfunction.imcsch.md)|
|[ImDiv](Excel.WorksheetFunction.ImDiv.md)|
|[ImExp](Excel.WorksheetFunction.ImExp.md)|
|[ImLn](Excel.WorksheetFunction.ImLn.md)|
|[ImLog10](Excel.WorksheetFunction.ImLog10.md)|
|[ImLog2](Excel.WorksheetFunction.ImLog2.md)|
|[ImPower](Excel.WorksheetFunction.ImPower.md)|
|[ImProduct](Excel.WorksheetFunction.ImProduct.md)|
|[ImReal](Excel.WorksheetFunction.ImReal.md)|
|[ImSec](Excel.worksheetfunction.imsec.md)|
|[ImSech](Excel.worksheetfunction.imsech.md)|
|[ImSin](Excel.WorksheetFunction.ImSin.md)|
|[ImSinh](Excel.worksheetfunction.imsinh.md)|
|[ImSqrt](Excel.WorksheetFunction.ImSqrt.md)|
|[ImSub](Excel.WorksheetFunction.ImSub.md)|
|[ImSum](Excel.WorksheetFunction.ImSum.md)|
|[ImTan](Excel.worksheetfunction.imtan.md)|
|[Index](Excel.WorksheetFunction.Index.md)|
|[Intercept](Excel.WorksheetFunction.Intercept.md)|
|[IntRate](Excel.WorksheetFunction.IntRate.md)|
|[Ipmt](Excel.WorksheetFunction.Ipmt.md)|
|[Irr](Excel.WorksheetFunction.Irr.md)|
|[IsErr](Excel.WorksheetFunction.IsErr.md)|
|[IsError](Excel.WorksheetFunction.IsError.md)|
|[IsEven](Excel.WorksheetFunction.IsEven.md)|
|[IsFormula](Excel.worksheetfunction.isformula.md)|
|[IsLogical](Excel.WorksheetFunction.IsLogical.md)|
|[IsNA](Excel.WorksheetFunction.IsNA.md)|
|[IsNonText](Excel.WorksheetFunction.IsNonText.md)|
|[IsNumber](Excel.WorksheetFunction.IsNumber.md)|
|[ISO_Ceiling](Excel.WorksheetFunction.ISO_Ceiling.md)|
|[IsOdd](Excel.WorksheetFunction.IsOdd.md)|
|[IsoWeekNum](Excel.worksheetfunction.isoweeknum.md)|
|[Ispmt](Excel.WorksheetFunction.Ispmt.md)|
|[IsText](Excel.WorksheetFunction.IsText.md)|
|[Kurt](Excel.WorksheetFunction.Kurt.md)|
|[Large](Excel.WorksheetFunction.Large.md)|
|[Lcm](Excel.WorksheetFunction.Lcm.md)|
|[LinEst](Excel.WorksheetFunction.LinEst.md)|
|[Ln](Excel.WorksheetFunction.Ln.md)|
|[Log](Excel.WorksheetFunction.Log.md)|
|[Log10](Excel.WorksheetFunction.Log10.md)|
|[LogEst](Excel.WorksheetFunction.LogEst.md)|
|[LogInv](Excel.WorksheetFunction.LogInv.md)|
|[LogNorm_Dist](Excel.WorksheetFunction.LogNorm_Dist.md)|
|[LogNorm_Inv](Excel.WorksheetFunction.LogNorm_Inv.md)|
|[LogNormDist](Excel.WorksheetFunction.LogNormDist.md)|
|[Lookup](Excel.WorksheetFunction.Lookup.md)|
|[Match](Excel.WorksheetFunction.Match.md)|
|[Max](Excel.WorksheetFunction.Max.md)|
|[MDeterm](Excel.WorksheetFunction.MDeterm.md)|
|[MDuration](Excel.WorksheetFunction.MDuration.md)|
|[Median](Excel.WorksheetFunction.Median.md)|
|[Min](Excel.WorksheetFunction.Min.md)|
|[MInverse](Excel.WorksheetFunction.MInverse.md)|
|[MIrr](Excel.WorksheetFunction.MIrr.md)|
|[MMult](Excel.WorksheetFunction.MMult.md)|
|[Mode](Excel.WorksheetFunction.Mode.md)|
|[Mode_Mult](Excel.WorksheetFunction.Mode_Mult.md)|
|[Mode_Sngl](Excel.WorksheetFunction.Mode_Sngl.md)|
|[MRound](Excel.WorksheetFunction.MRound.md)|
|[MultiNomial](Excel.WorksheetFunction.MultiNomial.md)|
|[Munit](Excel.worksheetfunction.munit.md)|
|[NegBinom_Dist](Excel.WorksheetFunction.NegBinom_Dist.md)|
|[NegBinomDist](Excel.WorksheetFunction.NegBinomDist.md)|
|[NetworkDays](Excel.WorksheetFunction.NetworkDays.md)|
|[NetworkDays_Intl](Excel.WorksheetFunction.NetworkDays_Intl.md)|
|[Nominal](Excel.WorksheetFunction.Nominal.md)|
|[Norm_Dist](Excel.WorksheetFunction.Norm_Dist.md)|
|[Norm_Inv](Excel.WorksheetFunction.Norm_Inv.md)|
|[Norm_S_Dist](Excel.WorksheetFunction.Norm_S_Dist.md)|
|[Norm_S_Inv](Excel.WorksheetFunction.Norm_S_Inv.md)|
|[NormDist](Excel.WorksheetFunction.NormDist.md)|
|[NormInv](Excel.WorksheetFunction.NormInv.md)|
|[NormSDist](Excel.WorksheetFunction.NormSDist.md)|
|[NormSInv](Excel.WorksheetFunction.NormSInv.md)|
|[NPer](Excel.WorksheetFunction.NPer.md)|
|[Npv](Excel.WorksheetFunction.Npv.md)|
|[NumberValue](Excel.worksheetfunction.numbervalue.md)|
|[Oct2Bin](Excel.WorksheetFunction.Oct2Bin.md)|
|[Oct2Dec](Excel.WorksheetFunction.Oct2Dec.md)|
|[Oct2Hex](Excel.WorksheetFunction.Oct2Hex.md)|
|[Odd](Excel.WorksheetFunction.Odd.md)|
|[OddFPrice](Excel.WorksheetFunction.OddFPrice.md)|
|[OddFYield](Excel.WorksheetFunction.OddFYield.md)|
|[OddLPrice](Excel.WorksheetFunction.OddLPrice.md)|
|[OddLYield](Excel.WorksheetFunction.OddLYield.md)|
|[Or](Excel.WorksheetFunction.Or.md)|
|[PDuration](Excel.worksheetfunction.pduration.md)|
|[Pearson](Excel.WorksheetFunction.Pearson.md)|
|[Percentile](Excel.WorksheetFunction.Percentile.md)|
|[Percentile_Exc](Excel.WorksheetFunction.Percentile_Exc.md)|
|[Percentile_Inc](Excel.WorksheetFunction.Percentile_Inc.md)|
|[PercentRank](Excel.WorksheetFunction.PercentRank.md)|
|[PercentRank_Exc](Excel.WorksheetFunction.PercentRank_Exc.md)|
|[PercentRank_Inc](Excel.WorksheetFunction.PercentRank_Inc.md)|
|[Permut](Excel.WorksheetFunction.Permut.md)|
|[Permutationa](Excel.worksheetfunction.permutationa.md)|
|[Phi](Excel.worksheetfunction.phi.md)|
|[Phonetic](Excel.WorksheetFunction.Phonetic.md)|
|[Pi](Excel.WorksheetFunction.Pi.md)|
|[Pmt](Excel.WorksheetFunction.Pmt.md)|
|[Poisson](Excel.WorksheetFunction.Poisson.md)|
|[Poisson_Dist](Excel.WorksheetFunction.Poisson_Dist.md)|
|[Power](Excel.WorksheetFunction.Power.md)|
|[Ppmt](Excel.WorksheetFunction.Ppmt.md)|
|[Price](Excel.WorksheetFunction.Price.md)|
|[PriceDisc](Excel.WorksheetFunction.PriceDisc.md)|
|[PriceMat](Excel.WorksheetFunction.PriceMat.md)|
|[Prob](Excel.WorksheetFunction.Prob.md)|
|[Product](Excel.WorksheetFunction.Product.md)|
|[Proper](Excel.WorksheetFunction.Proper.md)|
|[Pv](Excel.WorksheetFunction.Pv.md)|
|[Quartile](Excel.WorksheetFunction.Quartile.md)|
|[Quartile_Exc](Excel.WorksheetFunction.Quartile_Exc.md)|
|[Quartile_Inc](Excel.WorksheetFunction.Quartile_Inc.md)|
|[Quotient](Excel.WorksheetFunction.Quotient.md)|
|[Radians](Excel.WorksheetFunction.Radians.md)|
|[RandBetween](Excel.WorksheetFunction.RandBetween.md)|
|[Rank](Excel.WorksheetFunction.Rank.md)|
|[Rank_Avg](Excel.WorksheetFunction.Rank_Avg.md)|
|[Rank_Eq](Excel.WorksheetFunction.Rank_Eq.md)|
|[Rate](Excel.WorksheetFunction.Rate.md)|
|[Received](Excel.WorksheetFunction.Received.md)|
|[Replace](Excel.WorksheetFunction.Replace.md)|
|[ReplaceB](Excel.WorksheetFunction.ReplaceB.md)|
|[Rept](Excel.WorksheetFunction.Rept.md)|
|[Roman](Excel.WorksheetFunction.Roman.md)|
|[Round](Excel.WorksheetFunction.Round.md)|
|[RoundDown](Excel.WorksheetFunction.RoundDown.md)|
|[RoundUp](Excel.WorksheetFunction.RoundUp.md)|
|[Rri](Excel.worksheetfunction.rri.md)|
|[RSq](Excel.WorksheetFunction.RSq.md)|
|[RTD](Excel.WorksheetFunction.RTD.md)|
|[Search](Excel.WorksheetFunction.Search.md)|
|[SearchB](Excel.WorksheetFunction.SearchB.md)|
|[Sec](Excel.worksheetfunction.sec.md)|
|[Sech](Excel.worksheetfunction.sech.md)|
|[SeriesSum](Excel.WorksheetFunction.SeriesSum.md)|
|[Sinh](Excel.WorksheetFunction.Sinh.md)|
|[Skew](Excel.WorksheetFunction.Skew.md)|
|[Skew_p](Excel.worksheetfunction.skew_p.md)|
|[Sln](Excel.WorksheetFunction.Sln.md)|
|[Slope](Excel.WorksheetFunction.Slope.md)|
|[Small](Excel.WorksheetFunction.Small.md)|
|[SqrtPi](Excel.WorksheetFunction.SqrtPi.md)|
|[Standardize](Excel.WorksheetFunction.Standardize.md)|
|[StDev](Excel.WorksheetFunction.StDev.md)|
|[StDev_P](Excel.WorksheetFunction.StDev_P.md)|
|[StDev_S](Excel.WorksheetFunction.StDev_S.md)|
|[StDevP](Excel.WorksheetFunction.StDevP.md)|
|[StEyx](Excel.WorksheetFunction.StEyx.md)|
|[Substitute](Excel.WorksheetFunction.Substitute.md)|
|[Subtotal](Excel.WorksheetFunction.Subtotal.md)|
|[Sum](Excel.WorksheetFunction.Sum.md)|
|[SumIf](Excel.WorksheetFunction.SumIf.md)|
|[SumIfs](Excel.WorksheetFunction.SumIfs.md)|
|[SumProduct](Excel.WorksheetFunction.SumProduct.md)|
|[SumSq](Excel.WorksheetFunction.SumSq.md)|
|[SumX2MY2](Excel.WorksheetFunction.SumX2MY2.md)|
|[SumX2PY2](Excel.WorksheetFunction.SumX2PY2.md)|
|[SumXMY2](Excel.WorksheetFunction.SumXMY2.md)|
|[Syd](Excel.WorksheetFunction.Syd.md)|
|[T_Dist](Excel.WorksheetFunction.T_Dist.md)|
|[T_Dist_2T](Excel.WorksheetFunction.T_Dist_2T.md)|
|[T_Dist_RT](Excel.WorksheetFunction.T_Dist_RT.md)|
|[T_Inv](Excel.WorksheetFunction.T_Inv.md)|
|[T_Inv_2T](Excel.WorksheetFunction.T_Inv_2T.md)|
|[T_Test](Excel.WorksheetFunction.T_Test.md)|
|[Tanh](Excel.WorksheetFunction.Tanh.md)|
|[TBillEq](Excel.WorksheetFunction.TBillEq.md)|
|[TBillPrice](Excel.WorksheetFunction.TBillPrice.md)|
|[TBillYield](Excel.WorksheetFunction.TBillYield.md)|
|[TDist](Excel.WorksheetFunction.TDist.md)|
|[Text](Excel.WorksheetFunction.Text.md)|
|[TInv](Excel.WorksheetFunction.TInv.md)|
|[Transpose](Excel.WorksheetFunction.Transpose.md)|
|[Trend](Excel.WorksheetFunction.Trend.md)|
|[Trim](Excel.WorksheetFunction.Trim.md)|
|[TrimMean](Excel.WorksheetFunction.TrimMean.md)|
|[TTest](Excel.WorksheetFunction.TTest.md)|
|[Unichar](Excel.worksheetfunction.unichar.md)|
|[Unicode](Excel.worksheetfunction.unicode.md)|
|[USDollar](Excel.WorksheetFunction.USDollar.md)|
|[Var](Excel.WorksheetFunction.Var.md)|
|[Var_P](Excel.WorksheetFunction.Var_P.md)|
|[Var_S](Excel.WorksheetFunction.Var_S.md)|
|[VarP](Excel.WorksheetFunction.VarP.md)|
|[Vdb](Excel.WorksheetFunction.Vdb.md)|
|[VLookup](Excel.WorksheetFunction.VLookup.md)|
|[WebService](Excel.worksheetfunction.webservice.md)|
|[Weekday](Excel.WorksheetFunction.Weekday.md)|
|[WeekNum](Excel.WorksheetFunction.WeekNum.md)|
|[Weibull](Excel.WorksheetFunction.Weibull.md)|
|[Weibull_Dist](Excel.WorksheetFunction.Weibull_Dist.md)|
|[WorkDay](Excel.WorksheetFunction.WorkDay.md)|
|[WorkDay_Intl](Excel.WorksheetFunction.WorkDay_Intl.md)|
|[Xirr](Excel.WorksheetFunction.Xirr.md)|
|[Xnpv](Excel.WorksheetFunction.Xnpv.md)|
|[Xor](Excel.worksheetfunction.xor.md)|
|[YearFrac](Excel.WorksheetFunction.YearFrac.md)|
|[YieldDisc](Excel.WorksheetFunction.YieldDisc.md)|
|[YieldMat](Excel.WorksheetFunction.YieldMat.md)|
|[Z_Test](Excel.WorksheetFunction.Z_Test.md)|
|[ZTest](Excel.WorksheetFunction.ZTest.md)|
|[Forecast_ETS](Excel.worksheetfunction.forecast_ets.md)|
|[Forecast_ETS_ConfInt](Excel.worksheetfunction.forecast_ets_confint.md)|
|[Forecast_ETS_Seasonality](Excel.worksheetfunction.forecast_ets_seasonality.md)|
|[Forecast_ETS_STAT](Excel.worksheetfunction.forecast_ets_stat.md)|
|[Forecast_Linear](Excel.worksheetfunction.forecast_linear.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.WorksheetFunction.Application.md)|
|[Creator](Excel.WorksheetFunction.Creator.md)|
|[Parent](Excel.WorksheetFunction.Parent.md)|

## See also

[Using a worksheet function in a Visual Basic macro in Excel](https://support.microsoft.com/help/291309/using-a-worksheet-function-in-a-visual-basic-macro-in-excel)
[Excel Object Model Reference](./overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]