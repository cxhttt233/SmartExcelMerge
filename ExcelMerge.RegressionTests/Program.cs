using System;
using System.Linq;
namespace ExcelMerge.RegressionTests
{
 internal static class Program
 {
  static int failures;
  static int Main(){ Run("auto",Auto); Run("manual single",Single); Run("composite",Composite); Run("analysis",Analysis); Run("invalid fallback",Invalid); Console.WriteLine(failures==0?"All regression tests passed.":failures+" failed."); return failures==0?0:1; }
  static ExcelSheet Sheet(params string[][] v){var s=new ExcelSheet(); for(int r=0;r<v.Length;r++)s.Rows.Add(r,new ExcelRow(r,v[r].Select((x,c)=>new ExcelCell(x,c,r)))); return s;}
  static ExcelSheetDiffConfig Config(){return new ExcelSheetDiffConfig{UseSmartTableDiff=true,SrcHeaderIndex=0,DstHeaderIndex=0,SrcSheetIndex=0,DstSheetIndex=0};}
  static void Auto(){RowKeySelectionRuntime.Clear();var c=Config();var d=ExcelSheet.Diff(Sheet(new[]{"ID","Name"},new[]{"1","A"},new[]{"2","B"}),Sheet(new[]{"ID","Name"},new[]{"2","B2"},new[]{"1","A"}),c).CreateSummary(); A(d.ModifiedRowCount==1&&d.AddedRowCount==0&&d.RemovedRowCount==0,"auto pairing"); A(c.RowKeyAnalysis.SelectedDisplayName=="ID","auto ID");}
  static void Single(){RowKeySelectionRuntime.Clear();RowKeySelectionRuntime.SetManualSelection(0,0,new[]{"Code"});var c=Config();var d=ExcelSheet.Diff(Sheet(new[]{"Seq","Code","Name"},new[]{"1","A1","A"},new[]{"2","A2","B"}),Sheet(new[]{"Seq","Code","Name"},new[]{"2","A2","B2"},new[]{"1","A1","A"}),c).CreateSummary();A(d.AddedRowCount==0&&d.RemovedRowCount==0,"manual pairing");A(c.RowKeyAnalysis.SelectionMode==RowKeySelectionMode.Manual,"manual mode");}
  static void Composite(){RowKeySelectionRuntime.Clear();RowKeySelectionRuntime.SetManualSelection(0,0,new[]{"Region","Name"});var c=Config();var d=ExcelSheet.Diff(Sheet(new[]{"Region","Name","Value"},new[]{"N","Gate","10"},new[]{"S","Gate","20"}),Sheet(new[]{"Region","Name","Value"},new[]{"S","Gate","21"},new[]{"N","Gate","10"}),c).CreateSummary();A(d.ModifiedRowCount==1&&d.AddedRowCount==0&&d.RemovedRowCount==0,"composite pairing");A(c.RowKeyAnalysis.SelectedColumnNames.Count==2,"composite selected");}
  static void Analysis(){RowKeySelectionRuntime.Clear();var c=Config();ExcelSheet.Diff(Sheet(new[]{"ID","Name"},new[]{"1","A"},new[]{"2","B"}),Sheet(new[]{"ID","Name"},new[]{"1","A"},new[]{"2","B2"}),c);A(c.RowKeyAnalysis.Candidates.Any(x=>x.DisplayName=="ID"),"candidate");A(!string.IsNullOrWhiteSpace(c.RowKeyAnalysis.SelectionReason),"reason");}
  static void Invalid(){RowKeySelectionRuntime.Clear();RowKeySelectionRuntime.SetManualSelection(0,0,new[]{"Missing"});var c=Config();ExcelSheet.Diff(Sheet(new[]{"ID"},new[]{"1"}),Sheet(new[]{"ID"},new[]{"1"}),c);A(c.RowKeyAnalysis!=null&&!c.RowKeyAnalysis.HasSelectedKey,"fallback");}
  static void Run(string n,Action t){try{t();Console.WriteLine("[PASS] "+n);}catch(Exception e){failures++;Console.WriteLine("[FAIL] "+n+": "+e.Message);}}
  static void A(bool c,string m){if(!c)throw new InvalidOperationException(m);}
 }
}
