using System;
using System.Collections.Generic;

namespace Modelwright.Core.Trace;

/// <summary>
/// What Trace In's evaluate mode needs to know about a function before it evaluates a fragment that calls it
/// (docs/PLAN.md section 4.5): whether it is one of Excel's own worksheet functions, whether it reaches outside the
/// workbook (a web service, a data feed, an OLAP cube, a DLL), and whether it is volatile (a new value at every
/// evaluation). A function that is not Excel's (a VBA or add-in function, an external workbook's function, a
/// LAMBDA held by a defined name) may have side effects, so it is never evaluated.
/// </summary>
/// <remarks>
/// Names are matched as <see cref="FunctionSignatures.Normalize"/> writes them (upper case, without <c>_xlfn.</c> /
/// <c>_xlws.</c>). Source: Microsoft's Excel function reference (every worksheet function, the newest included).
/// </remarks>
public static class ExcelFunctions
{
    private static readonly HashSet<string> BuiltIn = Set(
        "ABS ACCRINT ACCRINTM ACOS ACOSH ACOT ACOTH ADDRESS AGGREGATE AMORDEGRC AMORLINC AND ANCHORARRAY ARABIC AREAS " +
        "ARRAYTOTEXT ASC ASIN ASINH ATAN ATAN2 ATANH AVEDEV AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS BAHTTEXT BASE BESSELI " +
        "BESSELJ BESSELK BESSELY BETA.DIST BETA.INV BETADIST BETAINV BIN2DEC BIN2HEX BIN2OCT BINOM.DIST BINOM.DIST.RANGE " +
        "BINOM.INV BINOMDIST BITAND BITLSHIFT BITOR BITRSHIFT BITXOR BYCOL BYROW CALL CEILING CEILING.MATH " +
        "CEILING.PRECISE CELL CHAR CHIDIST CHIINV CHISQ.DIST CHISQ.DIST.RT CHISQ.INV CHISQ.INV.RT CHISQ.TEST CHITEST " +
        "CHOOSE CHOOSECOLS CHOOSEROWS CLEAN CODE COLUMN COLUMNS COMBIN COMBINA COMPLEX CONCAT CONCATENATE CONFIDENCE " +
        "CONFIDENCE.NORM CONFIDENCE.T CONVERT COPILOT CORREL COS COSH COT COTH COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS " +
        "COUPDAYBS COUPDAYS COUPDAYSNC COUPNCD COUPNUM COUPPCD COVAR COVARIANCE.P COVARIANCE.S CRITBINOM CSC CSCH " +
        "CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE CUMIPMT CUMPRINC " +
        "DATE DATEDIF DATEVALUE DAVERAGE DAY DAYS DAYS360 DB DBCS DCOUNT DCOUNTA DDB DEC2BIN DEC2HEX DEC2OCT DECIMAL " +
        "DEGREES DELTA DETECTLANGUAGE DEVSQ DGET DISC DMAX DMIN DOLLAR DOLLARDE DOLLARFR DPRODUCT DROP DSTDEV DSTDEVP " +
        "DSUM DURATION DVAR DVARP ECMA.CEILING EDATE EFFECT ENCODEURL EOMONTH ERF ERF.PRECISE ERFC ERFC.PRECISE " +
        "ERROR.TYPE EUROCONVERT EVEN EXACT EXP EXPAND EXPON.DIST EXPONDIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FACT " +
        "FACTDOUBLE FALSE FDIST FIELDVALUE FILTER FILTERXML FIND FINDB FINV FISHER FISHERINV FIXED FLOOR FLOOR.MATH " +
        "FLOOR.PRECISE FORECAST FORECAST.ETS FORECAST.ETS.CONFINT FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT " +
        "FORECAST.LINEAR FORMULATEXT FREQUENCY FTEST FV FVSCHEDULE GAMMA GAMMA.DIST GAMMA.INV GAMMADIST GAMMAINV GAMMALN " +
        "GAMMALN.PRECISE GAUSS GCD GEOMEAN GESTEP GETPIVOTDATA GROUPBY GROWTH HARMEAN HEX2BIN HEX2DEC HEX2OCT HLOOKUP " +
        "HOUR HSTACK HYPERLINK HYPGEOM.DIST HYPGEOMDIST IF IFERROR IFNA IFS IMABS IMAGE IMAGINARY IMARGUMENT IMCONJUGATE " +
        "IMCOS IMCOSH IMCOT IMCSC IMCSCH IMDIV IMEXP IMLN IMLOG10 IMLOG2 IMPOWER IMPRODUCT IMREAL IMSEC IMSECH IMSIN " +
        "IMSINH IMSQRT IMSUB IMSUM IMTAN INDEX INDIRECT INFO INT INTERCEPT INTRATE IPMT IRR ISBLANK ISERR ISERROR ISEVEN " +
        "ISFORMULA ISLOGICAL ISNA ISNONTEXT ISNUMBER ISO.CEILING ISODD ISOMITTED ISOWEEKNUM ISPMT ISREF ISTEXT JIS KURT " +
        "LAMBDA LARGE LCM LEFT LEFTB LEN LENB LET LINEST LN LOG LOG10 LOGEST LOGINV LOGNORM.DIST LOGNORM.INV LOGNORMDIST " +
        "LOOKUP LOWER MAKEARRAY MAP MATCH MAX MAXA MAXIFS MDETERM MDURATION MEDIAN MID MIDB MIN MINA MINIFS MINUTE " +
        "MINVERSE MIRR MMULT MOD MODE MODE.MULT MODE.SNGL MONTH MROUND MULTINOMIAL MUNIT N NA NEGBINOM.DIST NEGBINOMDIST " +
        "NETWORKDAYS NETWORKDAYS.INTL NOMINAL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV NORMDIST NORMINV NORMSDIST " +
        "NORMSINV NOT NOW NPER NPV NUMBERVALUE OCT2BIN OCT2DEC OCT2HEX ODD ODDFPRICE ODDFYIELD ODDLPRICE ODDLYIELD OFFSET " +
        "OR PDURATION PEARSON PERCENTILE PERCENTILE.EXC PERCENTILE.INC PERCENTOF PERCENTRANK PERCENTRANK.EXC " +
        "PERCENTRANK.INC PERMUT PERMUTATIONA PHI PHONETIC PI PIVOTBY PMT POISSON POISSON.DIST POWER PPMT PRICE PRICEDISC " +
        "PRICEMAT PROB PRODUCT PROPER PV PY QUARTILE QUARTILE.EXC QUARTILE.INC QUOTIENT RADIANS RAND RANDARRAY " +
        "RANDBETWEEN RANK RANK.AVG RANK.EQ RATE RECEIVED REDUCE REGEXEXTRACT REGEXREPLACE REGEXTEST REGISTER.ID REPLACE " +
        "REPLACEB REPT RIGHT RIGHTB ROMAN ROUND ROUNDDOWN ROUNDUP ROW ROWS RRI RSQ RTD SCAN SEARCH SEARCHB SEC SECH " +
        "SECOND SEQUENCE SERIESSUM SHEET SHEETS SIGN SIN SINGLE SINH SKEW SKEW.P SLN SLOPE SMALL SORT SORTBY SQRT SQRTPI " +
        "STANDARDIZE STDEV STDEV.P STDEV.S STDEVA STDEVP STDEVPA STEYX STOCKHISTORY SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS " +
        "SUMPRODUCT SUMSQ SUMX2MY2 SUMX2PY2 SUMXMY2 SWITCH SYD T T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST TAKE " +
        "TAN TANH TBILLEQ TBILLPRICE TBILLYIELD TDIST TEXT TEXTAFTER TEXTBEFORE TEXTJOIN TEXTSPLIT TIME TIMEVALUE TINV " +
        "TOCOL TODAY TOROW TRANSLATE TRANSPOSE TREND TRIM TRIMMEAN TRIMRANGE TRUE TRUNC TTEST TYPE UNICHAR UNICODE " +
        "UNIQUE UPPER VALUE VALUETOTEXT VAR VAR.P VAR.S VARA VARP VARPA VDB VLOOKUP VSTACK WEBSERVICE WEEKDAY WEEKNUM " +
        "WEIBULL WEIBULL.DIST WORKDAY WORKDAY.INTL WRAPCOLS WRAPROWS XIRR XLOOKUP XMATCH XNPV XOR YEAR YEARFRAC YIELD " +
        "YIELDDISC YIELDMAT Z.TEST ZTEST");

    // Excel's functions that reach outside the workbook (the web, a data feed, an OLAP cube, a cloud service) or run
    // code in a DLL: evaluating one per row would wait on the network or act, so they are never evaluated.
    private static readonly HashSet<string> Outside = Set(
        "CALL COPILOT CUBEKPIMEMBER CUBEMEMBER CUBEMEMBERPROPERTY CUBERANKEDMEMBER CUBESET CUBESETCOUNT CUBEVALUE " +
        "DETECTLANGUAGE IMAGE PY REGISTER.ID RTD STOCKHISTORY TRANSLATE WEBSERVICE");

    // A new value at every evaluation: the row's value can differ from the cell's.
    private static readonly HashSet<string> VolatileFunctions = Set("NOW RAND RANDARRAY RANDBETWEEN TODAY");

    /// <summary>
    /// True for one of Excel's worksheet functions (or one in <see cref="FunctionSignatures"/>); false for anything
    /// else: a VBA or add-in function, a function in another workbook (<c>Book.xlsx!MyUdf</c>), a LAMBDA held by a
    /// defined name or a LET name (<c>_xlpm.f</c>).
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static bool IsBuiltIn(string functionName)
    {
        var name = FunctionSignatures.Normalize(functionName);
        return BuiltIn.Contains(name) || FunctionSignatures.Contains(name);
    }

    /// <summary>
    /// True for one of Excel's functions that reach outside the workbook or run code: WEBSERVICE, STOCKHISTORY, RTD,
    /// the CUBE functions, IMAGE, PY, COPILOT, TRANSLATE, DETECTLANGUAGE, CALL and REGISTER.ID.
    /// </summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static bool ReachesOutside(string functionName) => Outside.Contains(FunctionSignatures.Normalize(functionName));

    /// <summary>True for a volatile function whose value changes at every evaluation: NOW, TODAY, RAND, RANDBETWEEN, RANDARRAY.</summary>
    /// <exception cref="ArgumentNullException"><paramref name="functionName"/> is null.</exception>
    public static bool IsVolatile(string functionName) => VolatileFunctions.Contains(FunctionSignatures.Normalize(functionName));

    private static HashSet<string> Set(string names) =>
        new HashSet<string>(names.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), StringComparer.Ordinal);
}
