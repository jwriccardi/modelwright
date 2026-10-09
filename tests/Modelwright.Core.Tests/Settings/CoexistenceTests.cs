using Modelwright.Core.Settings;
using Xunit;

namespace Modelwright.Core.Tests.Settings;

public class CoexistenceTests
{
    [Theory]
    // Macabacus detected, shortcuts on, notice shown before, notices suppressed -> actions.
    [InlineData(false, true, false, false, CoexistenceActions.None)]
    [InlineData(false, false, false, false, CoexistenceActions.None)]
    [InlineData(false, true, true, true, CoexistenceActions.None)]
    [InlineData(true, false, false, false, CoexistenceActions.None)]
    [InlineData(true, false, true, false, CoexistenceActions.None)]
    [InlineData(true, false, false, true, CoexistenceActions.None)]
    [InlineData(true, true, false, false, CoexistenceActions.Reregister | CoexistenceActions.ShowNotice)]
    [InlineData(true, true, true, false, CoexistenceActions.Reregister)]
    [InlineData(true, true, false, true, CoexistenceActions.Reregister)]
    [InlineData(true, true, true, true, CoexistenceActions.Reregister)]
    public void Decide_follows_the_table(bool detected, bool shortcutsOn, bool noticeShown, bool suppressed, CoexistenceActions expected)
    {
        Assert.Equal(expected, Coexistence.Decide(detected, shortcutsOn, noticeShown, suppressed));
    }

    [Theory]
    [InlineData("1", true)]
    [InlineData(" 1 ", true)]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("true", false)]
    [InlineData("yes", false)]
    public void Notices_are_suppressed_only_by_1(string? value, bool expected)
    {
        Assert.Equal(expected, Coexistence.NoticesSuppressed(value));
        Assert.Equal("MODELWRIGHT_NO_NOTICES", Coexistence.NoNoticesVariable);
    }

    [Theory]
    [InlineData("Macabacus", true, true)]
    [InlineData("macabacus", true, true)]
    [InlineData(" Macabacus ", true, true)]
    [InlineData("Macabacus.AddIn", true, true)]
    [InlineData("Macabacus", false, false)]
    [InlineData("MacabacusTools", true, false)]
    [InlineData("Macabacus.xlam", false, false)]
    [InlineData("OtherVendor.Macabacus", true, false)]
    [InlineData("", true, false)]
    [InlineData(null, true, false)]
    public void Com_add_in_is_macabacus_when_its_progid_is_macabacus_and_it_is_connected(string? progId, bool connected, bool expected)
    {
        Assert.Equal(expected, Coexistence.IsMacabacusComAddIn(progId, connected));
    }
}
