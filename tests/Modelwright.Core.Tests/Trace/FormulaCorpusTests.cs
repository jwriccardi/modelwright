using System.Linq;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

/// <summary>
/// Real-world-style modeling formulas (DCF, LBO, three-statement links, lookups, names, tables, external links,
/// LET/LAMBDA, dynamic arrays, nasty quoting) and the references each must yield, in order. The formulas live on
/// sheet <c>Calc</c> of <c>Model.xlsx</c>, so references to that sheet or workbook come back unqualified.
/// </summary>
/// <remarks>
/// Expected references are written as <c>Kind location+target</c> (see <see cref="Describe"/>), separated by
/// <c>"; "</c>: <c>Cell A1</c>, <c>Range Debt!E5:E9</c>, <c>Col A:C</c>, <c>Row 3:5</c>, <c>Name Inputs!TaxRate</c>,
/// <c>Local x</c>, <c>Table Sales[@][Amount]</c> (<c>?</c> for no table name), <c>RefErr #REF!</c>; an external
/// workbook is <c>path[Book.xlsx]</c>, a 3-D sheet span <c>First:Last</c>, a spill a trailing <c>#</c>.
/// </remarks>
public class FormulaCorpusTests
{
    private static readonly FormulaContext Context = new FormulaContext("Model.xlsx", "Calc");

    public static TheoryData<string, string> Corpus => new TheoryData<string, string>
    {
        // DCF
        { "=SUM(D10:D14)", "Range D10:D14" },
        { "=D20*(1-$C$5)", "Cell D20; Cell C5" },
        { "=E25/(1+$C$8)^E$3", "Cell E25; Cell C8; Cell E3" },
        { "=NPV($C$8,E25:I25)", "Cell C8; Range E25:I25" },
        { "=I25*(1+$C$9)/($C$8-$C$9)", "Cell I25; Cell C9; Cell C8; Cell C9" },
        { "=SUM(E27:I27)+I29/(1+$C$8)^I$3", "Range E27:I27; Cell I29; Cell C8; Cell I3" },
        { "=Assumptions!$C$12*E15", "Cell Assumptions!C12; Cell E15" },
        { "=IFERROR(E30/E31,0)", "Cell E30; Cell E31" },
        { "=XNPV(WACC,E25:I25,E4:I4)", "Name WACC; Range E25:I25; Range E4:I4" },
        { "=XIRR(E40:I40,E4:I4)", "Range E40:I40; Range E4:I4" },
        { "=IRR(C40:I40)", "Range C40:I40" },
        { "=E12-E13-E14", "Cell E12; Cell E13; Cell E14" },
        { "=EBITDA*Multiple", "Name EBITDA; Name Multiple" },
        { "=Calc!E5+E6", "Cell E5; Cell E6" },
        { "='Calc'!E5", "Cell E5" },
        { "=MAX(0,E50-E51)", "Cell E50; Cell E51" },
        { "=MIN(E60,$C$62*E61)", "Cell E60; Cell C62; Cell E61" },
        { "=ROUND(E70*'Tax Assumptions'!$D$4,0)", "Cell E70; Cell Tax Assumptions!D4" },
        { "=(1+$C$7)^(1/12)-1", "Cell C7" },
        { "=PMT($C$15/12,$C$16*12,-$C$14)", "Cell C15; Cell C16; Cell C14" },
        { "=PV(Rate,Nper,,-FV_Amount)", "Name Rate; Name Nper; Name FV_Amount" },
        { "=FV(C5,C6,C7,C8,0)", "Cell C5; Cell C6; Cell C7; Cell C8" },
        { "=EOMONTH(E$4,12)", "Cell E4" },
        { "=EDATE(StartDate,3*(COLUMN()-COLUMN($E$4)))", "Name StartDate; Cell E4" },
        { "=YEAR(E4)&\"E\"", "Cell E4" },
        { "=DATE(YEAR(E4),MONTH(E4)+3,1)", "Cell E4; Cell E4" },
        { "=YEARFRAC($C$4,E$4,1)", "Cell C4; Cell E4" },
        { "=E40*(1+Growth)", "Cell E40; Name Growth" },
        { "=SUMPRODUCT(E25:I25,E26:I26)", "Range E25:I25; Range E26:I26" },
        { "=AVERAGE(Comps!$D$5:$D$15)", "Range Comps!D5:D15" },
        { "=MEDIAN(Comps!D5:D15)*$C$30", "Range Comps!D5:D15; Cell C30" },
        { "=E33*E$34", "Cell E33; Cell E34" },

        // LBO
        { "=Debt!E15+Debt!E25+Debt!E35", "Cell Debt!E15; Cell Debt!E25; Cell Debt!E35" },
        { "=-MIN(E50,SUM(E45:E49))", "Cell E50; Range E45:E49" },
        { "=MAX(MIN(E60,E61-E62),0)", "Cell E60; Cell E61; Cell E62" },
        { "=IF(E$3<=$C$20,E70*$C$21,0)", "Cell E3; Cell C20; Cell E70; Cell C21" },
        { "=IF(AND(E$3>=$C$5,E$3<=$C$6),1,0)", "Cell E3; Cell C5; Cell E3; Cell C6" },
        { "=D80+E78-E79", "Cell D80; Cell E78; Cell E79" },
        { "=AVERAGE(D80:E80)*InterestRate", "Range D80:E80; Name InterestRate" },
        { "=E90*(1-'Sources & Uses'!$C$10)", "Cell E90; Cell Sources & Uses!C10" },
        { "='Sources & Uses'!C20/'Sources & Uses'!C25", "Cell Sources & Uses!C20; Cell Sources & Uses!C25" },
        { "=(I100+I101)/(Entry_Equity)", "Cell I100; Cell I101; Name Entry_Equity" },
        { "=(I105/C105)^(1/(I3-C3))-1", "Cell I105; Cell C105; Cell I3; Cell C3" },
        { "=IFERROR(I110/I111,\"n.m.\")", "Cell I110; Cell I111" },
        { "=SUM(OFFSET(E50,0,0,1,$C$4))", "Cell E50; Cell C4" },
        { "=INDEX(E50:I50,1,MATCH(ExitYear,E3:I3,0))", "Range E50:I50; Name ExitYear; Range E3:I3" },
        { "=CHOOSE(Scenario,E10,E11,E12)", "Name Scenario; Cell E10; Cell E11; Cell E12" },
        { "=CHOOSE($C$3,Base!E10,Upside!E10,Downside!E10)", "Cell C3; Cell Base!E10; Cell Upside!E10; Cell Downside!E10" },

        // Three-statement links
        { "=IS!E25", "Cell IS!E25" },
        { "='Income Statement'!E25-'Income Statement'!E30", "Cell Income Statement!E25; Cell Income Statement!E30" },
        { "=BS!D40+CF!E55", "Cell BS!D40; Cell CF!E55" },
        { "=E10-E11+E12-SUM(E13:E15)", "Cell E10; Cell E11; Cell E12; Range E13:E15" },
        { "=-('Balance Sheet'!E12-'Balance Sheet'!D12)", "Cell Balance Sheet!E12; Cell Balance Sheet!D12" },
        { "=SUM('Balance Sheet'!E5:E12)-SUM('Balance Sheet'!E20:E30)", "Range Balance Sheet!E5:E12; Range Balance Sheet!E20:E30" },
        { "=ROUND(BS!E50-BS!E80,2)=0", "Cell BS!E50; Cell BS!E80" },
        { "=IF(ABS(Check)>0.001,\"ERROR\",\"OK\")", "Name Check" },
        { "=E5*DSO/365", "Cell E5; Name DSO" },
        { "=Revenue_2024*(1+Rev_Growth)", "Name Revenue_2024; Name Rev_Growth" },
        { "=SUM(E5:E9 E7:I7)", "Range E5:E9; Range E7:I7" },
        { "=SUM((E5:E9,G5:G9))", "Range E5:E9; Range G5:G9" },
        { "=SUM(E5:E9,G5:G9)", "Range E5:E9; Range G5:G9" },

        // 3-D references
        { "=SUM('Q1:Q4'!E5)", "Cell Q1:Q4!E5" },
        { "=SUM(Jan:Dec!B5)", "Cell Jan:Dec!B5" },
        { "=SUM('Jan 2024:Dec 2024'!B5:B10)", "Range Jan 2024:Dec 2024!B5:B10" },
        { "=AVERAGE(Sheet1:Sheet3!A:A)", "Col Sheet1:Sheet3!A:A" },
        { "=SUM('It''s:Other'!A1)", "Cell It's:Other!A1" },

        // Duplicates, absolute and mixed references, normalization
        { "=E5-E5", "Cell E5; Cell E5" },
        { "=E5+E5+E5", "Cell E5; Cell E5; Cell E5" },
        { "=$E$5+E$5+$E5+E5", "Cell E5; Cell E5; Cell E5; Cell E5" },
        { "=SUM(B5:B1)", "Range B1:B5" },
        { "=SUM(D3:B1)", "Range B1:D3" },
        { "=SUM(A:A)", "Col A:A" },
        { "=SUM($A:$C)", "Col A:C" },
        { "=SUM(C:A)", "Col A:C" },
        { "=SUM(1:1)", "Row 1:1" },
        { "=SUM($5:$3)", "Row 3:5" },
        { "=COUNTA(Data!A:A)-1", "Col Data!A:A" },
        { "=SUM(A1:XFD1048576)", "Range A1:XFD1048576" },
        { "=XFD1048576", "Cell XFD1048576" },
        { "=TAX2023*2", "Cell TAX2023" },
        { "=FY2024+FY2025", "Cell FY2024; Cell FY2025" },
        { "=Rev_FY24*Margin.Target", "Name Rev_FY24; Name Margin.Target" },
        { "=_FY2024", "Name _FY2024" },
        { "=\\MyName", "Name \\MyName" },
        { "=A1B2", "Name A1B2" },

        // Defined names
        { "=Calc!LocalRate*E5", "Name LocalRate; Cell E5" },
        { "=Inputs!TaxRate*E5", "Name Inputs!TaxRate; Cell E5" },
        { "='Model Inputs'!TaxRate", "Name Model Inputs!TaxRate" },
        { "=SUM(Revenue)", "Name Revenue" },
        { "=VLOOKUP($B5,LookupTable,3,FALSE)", "Cell B5; Name LookupTable" },
        { "=INDEX(Prices,MATCH(Item,Items,0))", "Name Prices; Name Item; Name Items" },
        { "=Book2.xlsx!Rate", "Name [Book2.xlsx]Rate" },
        { "='C:\\Models\\Rates.xlsx'!BaseRate", "Name C:\\Models\\[Rates.xlsx]BaseRate" },
        { "=[Rates.xlsx]Curve!Spot1Y", "Name [Rates.xlsx]Curve!Spot1Y" },
        { "=Model.xlsx!WACC", "Name WACC" },
        { "=[Model.xlsx]Calc!E5", "Cell E5" },
        { "=Data.ods!Rate", "Name [Data.ods]Rate" },
        { "=Addin.xlam!Rate+Old.xls!Rate", "Name [Addin.xlam]Rate; Name [Old.xls]Rate" },
        { "='C:\\Deals [2024]\\Book.xlsx'!Rate", "Name C:\\Deals [2024]\\[Book.xlsx]Rate" },

        // Structured references
        { "=SUM(Sales[Amount])", "Table Sales[Amount]" },
        { "=SUMIFS(Sales[Amount],Sales[Region],$B5,Sales[Year],C$4)", "Table Sales[Amount]; Table Sales[Region]; Cell B5; Table Sales[Year]; Cell C4" },
        { "=Sales[@Amount]*Sales[@Qty]", "Table Sales[@][Amount]; Table Sales[@][Qty]" },
        { "=[@Amount]*[@Qty]", "Table ?[@][Amount]; Table ?[@][Qty]" },
        { "=Sales[[#This Row],[Amount]]", "Table Sales[#This Row][Amount]" },
        { "=SUM(Sales[[#Data],[Q1]:[Q4]])", "Table Sales[#Data][Q1][Q4]" },
        { "=Sales[#Totals]", "Table Sales[#Totals]" },
        { "=Sales[[#Headers],[Amount]]", "Table Sales[#Headers][Amount]" },
        { "=ROWS(Sales[#All])", "Table Sales[#All]" },
        { "=SUM(Sales[[Unit Price]])", "Table Sales[Unit Price]" },
        { "=Sales[@[Unit Price]]*[@Qty]", "Table Sales[@][Unit Price]; Table ?[@][Qty]" },
        { "=SUM(Data[Col'[1']])", "Table Data[Col[1]]" },
        { "=COUNTIFS(Deals[Sector],$A5,Deals[Close Date],\">=\"&$B$2)", "Table Deals[Sector]; Cell A5; Table Deals[Close Date]; Cell B2" },
        { "=Book2.xlsx!Sales[Amount]", "Table [Book2.xlsx]Sales[Amount]" },
        { "=SUBTOTAL(109,Sales[Amount])", "Table Sales[Amount]" },

        // External workbooks
        { "=[Budget.xlsx]Summary!$C$5", "Cell [Budget.xlsx]Summary!C5" },
        { "='[Budget 2025.xlsx]P&L'!C5:C10", "Range [Budget 2025.xlsx]P&L!C5:C10" },
        { "='C:\\Finance\\Models\\[Budget.xlsx]Summary'!$C$5", "Cell C:\\Finance\\Models\\[Budget.xlsx]Summary!C5" },
        { "='\\\\server\\share\\[Plan.xlsm]Inputs'!B2*2", "Cell \\\\server\\share\\[Plan.xlsm]Inputs!B2" },
        { "='https://contoso.sharepoint.com/sites/fin/Shared Documents/[Model.xlsx]Inputs'!$C$4", "Cell https://contoso.sharepoint.com/sites/fin/Shared Documents/[Model.xlsx]Inputs!C4" },
        { "=IFERROR(INDEX('[Comps.xlsx]Trading Comps'!$D:$D,MATCH($A5,'[Comps.xlsx]Trading Comps'!$A:$A,0)),\"n/a\")", "Col [Comps.xlsx]Trading Comps!D:D; Cell A5; Col [Comps.xlsx]Trading Comps!A:A" },
        { "='C:\\Users\\me\\[It''s.xlsx]Data'!B2", "Cell C:\\Users\\me\\[It's.xlsx]Data!B2" },
        { "=SUM([Budget.xlsx]Summary!A1:B2)+[Budget.xlsx]Summary!C3", "Range [Budget.xlsx]Summary!A1:B2; Cell [Budget.xlsx]Summary!C3" },
        { "=[Budget.xlsx]Summary!A:A", "Col [Budget.xlsx]Summary!A:A" },
        { "='[Book1.xlsx]Sheet1'!A1+[Book1.xlsx]Sheet1!A2", "Cell [Book1.xlsx]Sheet1!A1; Cell [Book1.xlsx]Sheet1!A2" },
        { "=VLOOKUP(A2,'[Price List.xlsx]Prices'!$A$2:$C$500,3,FALSE)", "Cell A2; Range [Price List.xlsx]Prices!A2:C500" },
        { "='C:\\Deals [2024]\\[Book.xlsx]Sheet1'!A1", "Cell C:\\Deals [2024]\\[Book.xlsx]Sheet1!A1" },
        { "=SUM('C:\\Deals [2024]\\[Book.xlsx]Jan:Dec'!B5)", "Cell C:\\Deals [2024]\\[Book.xlsx]Jan:Dec!B5" },
        { "='\\\\server\\share [x]\\[Plan.xlsm]Inputs'!B2", "Cell \\\\server\\share [x]\\[Plan.xlsm]Inputs!B2" },

        // Quoting, strings and constants that look like references
        { "='It''s'!A1", "Cell It's!A1" },
        { "='O''Brien''s Model'!B2:C3", "Range O'Brien's Model!B2:C3" },
        { "='Sheet (2)'!A1", "Cell Sheet (2)!A1" },
        { "='2024'!A1", "Cell 2024!A1" },
        { "='Q1-2024'!A1+'Q2-2024'!A1", "Cell Q1-2024!A1; Cell Q2-2024!A1" },
        { "='A1'!B2", "Cell A1!B2" },
        { "=Données!A1", "Cell Données!A1" },
        { "='Mgmt Case'!E10-'Bank Case'!E10", "Cell Mgmt Case!E10; Cell Bank Case!E10" },
        { "=\"Sheet1!A1\"&B1", "Cell B1" },
        { "=INDIRECT(\"'\"&$B$1&\"'!C5\")", "Cell B1" },
        { "=INDIRECT(ADDRESS(ROW(),COLUMN()-1))", "" },
        { "=IF(A1=\"Total\",SUM(B1:B9),\"\")", "Cell A1; Range B1:B9" },
        { "=TEXT(E4,\"mmm-yy\")", "Cell E4" },
        { "=TEXT(A1,\"$#,##0.00;($#,##0.00)\")", "Cell A1" },
        { "={1,2,3}", "" },
        { "=SUM({1,2,3}*A1:A3)", "Range A1:A3" },
        { "=MMULT(A1:B2,{1;2})", "Range A1:B2" },
        { "=IF(B5=\"A1:B2\",\"x\",C5)", "Cell B5; Cell C5" },
        { "={\"A1\",\"Sheet1!B2\";TRUE,#N/A}", "" },
        { "=E5&\" vs \"&F5", "Cell E5; Cell F5" },
        { "=--(A1>0)", "Cell A1" },
        { "=-A1%", "Cell A1" },
        { "=A1^2+B1^0.5", "Cell A1; Cell B1" },
        { "=A1<>B1", "Cell A1; Cell B1" },
        { "=NOT(ISBLANK(A1))", "Cell A1" },
        { "=#REF!+A1", "RefErr #REF!; Cell A1" },
        { "=Sheet9!#REF!*2", "RefErr Sheet9!#REF!" },
        { "=#REF!A1", "RefErr #REF!" },
        { "=SUM(#REF!B5:B10)*#REF!$C$3", "RefErr #REF!; RefErr #REF!" },
        { "=SUM(#REF!A:A)+A1", "RefErr #REF!; Cell A1" },
        { "=#REF!Rate*2", "RefErr #REF!" },
        { "=[Budget.xlsx]#REF!A1", "RefErr [Budget.xlsx]#REF!" },
        { "=SUM([Budget.xlsx]#REF!A1:B2)+C1", "RefErr [Budget.xlsx]#REF!; Cell C1" },
        { "=[Model.xlsx]#REF!A1", "RefErr #REF!" },
        { "='#REF'!A1", "Cell #REF!A1" },
        { "=SUM(A1:B2:C3)", "Range A1:C3" },
        { "=SUM(C3:A1:B2)", "Range A1:C3" },
        { "=SUM(A1:B2:C3:D4)", "Range A1:D4" },
        { "=SUM(Sheet2!A1:B2:Sheet2!C3)", "Range Sheet2!A1:C3" },
        { "=SUM(Sheet2!A1:B2:C3)", "Range Sheet2!A1:B2; Cell C3" },
        { "=SUM(A:B:C:D)", "Col A:B; Col C:D" },
        { "=SUM(A1,,B1)", "Cell A1; Cell B1" },
        { "=IF(A1,,)", "Cell A1" },
        { "=E5+\nE6", "Cell E5; Cell E6" },
        { "=SUM( E5 , E6 )", "Cell E5; Cell E6" },
        { "=IF(E5>0, E5 * $C$3, 0)", "Cell E5; Cell E5; Cell C3" },

        // LET and LAMBDA local names
        { "=LET(x,A1,y,B1,x+y)", "Local x; Cell A1; Local y; Cell B1; Local x; Local y" },
        { "=LET(rev,Revenue,cost,Cost,rev-cost)", "Local rev; Name Revenue; Local cost; Name Cost; Local rev; Local cost" },
        { "=LET(x,x,x*2)", "Local x; Name x; Local x" },
        { "=_xlfn.LET(_xlpm.x,A1,_xlpm.x*2)", "Local _xlpm.x; Cell A1; Local _xlpm.x" },
        { "=MAP(A1:A3,LAMBDA(v,v*2))", "Range A1:A3; Local v; Local v" },
        { "=BYROW(Data!A2:D100,LAMBDA(r,SUM(r)))", "Range Data!A2:D100; Local r; Local r" },
        { "=LET(rate,Inputs!C5,n,Inputs!C6,PMT(rate/12,n*12,-Loan))", "Local rate; Cell Inputs!C5; Local n; Cell Inputs!C6; Local rate; Local n; Name Loan" },
        { "=LET(a,1,b,a+Rate,b*a)", "Local a; Local b; Local a; Name Rate; Local b; Local a" },
        { "=LAMBDA(x,x+Rate)", "Local x; Local x; Name Rate" },
        { "=REDUCE(0,A1:A10,LAMBDA(acc,v,acc+v))", "Range A1:A10; Local acc; Local v; Local acc; Local v" },
        { "=x+LET(x,1,x)", "Name x; Local x; Local x" },
        { "=LET(total,SUM(A1:A10),IF(total>Limit,Limit,total))", "Local total; Range A1:A10; Local total; Name Limit; Name Limit; Local total" },
        { "=LET(X,A1,x*2)", "Local X; Cell A1; Local x" },

        // Dynamic arrays, spills and future functions
        { "=A1#", "Cell A1#" },
        { "=SUM(Calc2!B2#)", "Cell Calc2!B2#" },
        { "=FILTER(Data!A2:C100,Data!B2:B100>Threshold,\"none\")", "Range Data!A2:C100; Range Data!B2:B100; Name Threshold" },
        { "=SORT(UNIQUE(A2:A100))", "Range A2:A100" },
        { "=_xlfn._xlws.SORT(A2:A100,1,-1)", "Range A2:A100" },
        { "=SEQUENCE(12,1,StartDate,1)", "Name StartDate" },
        { "=XLOOKUP(A2,Map!A:A,Map!B:B,\"missing\",0)", "Cell A2; Col Map!A:A; Col Map!B:B" },
        { "=_xlfn.XLOOKUP(A2,B:B,C:C)", "Cell A2; Col B:B; Col C:C" },
        { "=@A1:A10", "Range A1:A10" },
        { "=@INDEX(A:A,ROW())", "Col A:A" },
        { "=HSTACK(A1:A3,B1:B3)", "Range A1:A3; Range B1:B3" },
        { "=TAKE(SORTBY(A2:C50,C2:C50,-1),5)", "Range A2:C50; Range C2:C50" },
        { "=SUM(A1:INDEX(B:B,5))", "Cell A1; Col B:B" },
        { "=INDEX(A1:C3,2,2):C5", "Range A1:C3; Cell C5" },
        { "=SUM(OFFSET(Data!$A$1,1,0,COUNTA(Data!$A:$A)-1,1))", "Cell Data!A1; Col Data!A:A" },
        { "=IFS(A1>90,\"A\",A1>80,\"B\",TRUE,\"C\")", "Cell A1; Cell A1" },
        { "=SWITCH(Region,\"N\",N_Rate,\"S\",S_Rate,Default_Rate)", "Name Region; Name N_Rate; Name S_Rate; Name Default_Rate" },
        { "=TEXTJOIN(\", \",TRUE,A1:A5)", "Range A1:A5" },
        { "=CONCAT(A1,\"-\",B1)", "Cell A1; Cell B1" },
        { "=MAXIFS(D:D,A:A,F2)", "Col D:D; Col A:A; Cell F2" },
        { "=NETWORKDAYS(A2,B2,Holidays)", "Cell A2; Cell B2; Name Holidays" },
        { "=MyUdf(A1,B2:C3)", "Cell A1; Range B2:C3" },
        { "=Book2.xlsx!MyUdf(A1)", "Cell A1" },
        { "=AGGREGATE(14,6,A1:A100/(B1:B100=\"x\"),1)", "Range A1:A100; Range B1:B100" },

        // Lookups
        { "=HLOOKUP(E$3,$E$40:$P$45,ROW()-ROW($A$40),FALSE)", "Cell E3; Range E40:P45; Cell A40" },
        { "=E5*(1+INDIRECT(\"Growth_\"&$C$2))", "Cell E5; Cell C2" },
        { "=IFERROR(VLOOKUP($B5,'Assumptions & Inputs'!$B$5:$Z$200,MATCH(F$3,'Assumptions & Inputs'!$B$4:$Z$4,0),FALSE),0)", "Cell B5; Range Assumptions & Inputs!B5:Z200; Cell F3; Range Assumptions & Inputs!B4:Z4" },
        { "=INDEX(Comps!$C$5:$H$30,MATCH($B7,Comps!$B$5:$B$30,0),MATCH(E$6,Comps!$C$4:$H$4,0))", "Range Comps!C5:H30; Cell B7; Range Comps!B5:B30; Cell E6; Range Comps!C4:H4" },
        { "=XLOOKUP($B5&\"|\"&C$4,Data!$A:$A&\"|\"&Data!$B:$B,Data!$C:$C,0)", "Cell B5; Cell C4; Col Data!A:A; Col Data!B:B; Col Data!C:C" },
        { "=SUMIFS(Data!$D:$D,Data!$A:$A,$B5,Data!$B:$B,\">=\"&E$3,Data!$B:$B,\"<=\"&EOMONTH(E$3,0))", "Col Data!D:D; Col Data!A:A; Cell B5; Col Data!B:B; Cell E3; Col Data!B:B; Cell E3" },
        { "=LOOKUP(2,1/(A1:A100<>\"\"),A1:A100)", "Range A1:A100; Range A1:A100" },
        { "=MATCH(TRUE,INDEX(E10:P10>0,0),0)", "Range E10:P10" },
        { "=COUNTIF(Checks!$C:$C,\"<>OK\")=0", "Col Checks!C:C" },
        { "=SUMPRODUCT((Deals[Sector]=$A5)*(Deals[Year]=C$4)*Deals[EV])", "Table Deals[Sector]; Cell A5; Table Deals[Year]; Cell C4; Table Deals[EV]" },
    };

    [Theory]
    [MemberData(nameof(Corpus))]
    public void Corpus_formula_yields_expected_references(string formula, string expected)
    {
        var parsed = FormulaParser.Parse(formula, Context);

        Assert.True(parsed.IsParsed, parsed.Error);
        Assert.Equal(expected, string.Join("; ", parsed.References.Select(Describe)));
        foreach (var reference in parsed.References)
        {
            Assert.Equal(reference.Text, formula.Substring(reference.Start, reference.Length));
        }

        var starts = parsed.References.Select(reference => reference.Start).ToList();
        Assert.Equal(starts.OrderBy(start => start), starts);
    }

    [Fact]
    public void Corpus_has_at_least_150_formulas()
    {
        Assert.True(Corpus.Count() >= 150, $"The corpus has {Corpus.Count()} formulas.");
    }

    internal static string Describe(FormulaReference reference)
    {
        var location = string.Empty;
        if (reference.WorkbookName is not null)
        {
            location += reference.WorkbookPath + "[" + reference.WorkbookName + "]";
        }

        if (reference.Sheet is not null)
        {
            location += reference.Sheet + (reference.LastSheet is null ? string.Empty : ":" + reference.LastSheet) + "!";
        }

        string kind;
        string target;
        switch (reference.Kind)
        {
            case FormulaReferenceKind.Cell:
                kind = "Cell";
                target = reference.Address!;
                break;
            case FormulaReferenceKind.Range:
                kind = "Range";
                target = reference.Address!;
                break;
            case FormulaReferenceKind.WholeColumn:
                kind = "Col";
                target = reference.Address!;
                break;
            case FormulaReferenceKind.WholeRow:
                kind = "Row";
                target = reference.Address!;
                break;
            case FormulaReferenceKind.Name:
                kind = "Name";
                target = reference.Name!;
                break;
            case FormulaReferenceKind.LocalName:
                kind = "Local";
                target = reference.Name!;
                break;
            case FormulaReferenceKind.StructuredReference:
                kind = "Table";
                target = (reference.Name ?? "?") +
                    string.Concat(reference.TableSpecifiers.Select(specifier => "[" + specifier + "]")) +
                    string.Concat(reference.TableColumns.Select(column => "[" + column + "]"));
                break;
            default:
                kind = "RefErr";
                target = "#REF!";
                break;
        }

        return kind + " " + location + target + (reference.IsSpill ? "#" : string.Empty);
    }
}
