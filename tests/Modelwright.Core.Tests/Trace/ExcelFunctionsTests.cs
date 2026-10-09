using System;
using Modelwright.Core.Trace;
using Xunit;

namespace Modelwright.Core.Tests.Trace;

public class ExcelFunctionsTests
{
    [Theory]
    [InlineData("SUM")]
    [InlineData("if")]
    [InlineData("INDEX")]
    [InlineData("_xlfn.XLOOKUP")]
    [InlineData("_xlfn._xlws.SORT")]
    [InlineData("TEXTSPLIT")]
    [InlineData("LAMBDA")]
    [InlineData("WEBSERVICE")]
    [InlineData("STDEV.S")]
    public void Excel_functions_are_built_in(string name)
    {
        Assert.True(ExcelFunctions.IsBuiltIn(name));
    }

    [Theory]
    [InlineData("MYUDF")]
    [InlineData("Book.xlsx!MyUdf")]
    [InlineData("_xlpm.f")]
    [InlineData("_xll.Price")]
    [InlineData("_xlfn.NOTAFUNCTION")]
    public void Anything_else_is_not(string name)
    {
        Assert.False(ExcelFunctions.IsBuiltIn(name));
    }

    [Fact]
    public void Every_function_with_a_signature_is_built_in()
    {
        Assert.All(FunctionSignatures.FunctionNames, name => Assert.True(ExcelFunctions.IsBuiltIn(name), name));
    }

    [Theory]
    [InlineData("WEBSERVICE", true)]
    [InlineData("_xlfn.STOCKHISTORY", true)]
    [InlineData("RTD", true)]
    [InlineData("CUBEVALUE", true)]
    [InlineData("CUBESET", true)]
    [InlineData("CALL", true)]
    [InlineData("SUM", false)]
    [InlineData("FILTERXML", false)]
    [InlineData("MYUDF", false)]
    public void Functions_that_reach_outside_are_known(string name, bool expected)
    {
        Assert.Equal(expected, ExcelFunctions.ReachesOutside(name));
    }

    [Theory]
    [InlineData("NOW", true)]
    [InlineData("TODAY", true)]
    [InlineData("RAND", true)]
    [InlineData("RANDBETWEEN", true)]
    [InlineData("_xlfn.RANDARRAY", true)]
    [InlineData("OFFSET", false)]
    [InlineData("SUM", false)]
    public void Volatile_functions_are_known(string name, bool expected)
    {
        Assert.Equal(expected, ExcelFunctions.IsVolatile(name));
    }

    [Fact]
    public void Null_names_are_rejected()
    {
        Assert.Throws<ArgumentNullException>(() => ExcelFunctions.IsBuiltIn(null!));
        Assert.Throws<ArgumentNullException>(() => ExcelFunctions.ReachesOutside(null!));
        Assert.Throws<ArgumentNullException>(() => ExcelFunctions.IsVolatile(null!));
    }
}
