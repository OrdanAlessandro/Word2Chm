using Word2Chm.Core.Common;

namespace Word2Chm.Core.Tests;

public sealed class ContextIdMarkerTests
{
    [Fact]
    public void ParsesSymbolOnly()
    {
        var marker = ContextIdMarker.Parse("Installazione {#IDH_INSTALLAZIONE}");

        Assert.Equal("IDH_INSTALLAZIONE", marker.Symbol);
        Assert.Null(marker.ExplicitId);
        Assert.Equal("Installazione", marker.CleanText);
    }

    [Fact]
    public void ParsesSymbolWithExplicitId()
    {
        var marker = ContextIdMarker.Parse("Riferimenti {#IDH_RIFERIMENTI=5000}");

        Assert.Equal("IDH_RIFERIMENTI", marker.Symbol);
        Assert.Equal(5000, marker.ExplicitId);
        Assert.Equal("Riferimenti", marker.CleanText);
    }

    [Fact]
    public void LeavesMarkedUpTextUntouchedWhenNoMarker()
    {
        var marker = ContextIdMarker.Parse("Testo normale");

        Assert.Equal(string.Empty, marker.Symbol);
        Assert.Equal("Testo normale", marker.CleanText);
    }

    [Fact]
    public void ToleratesSpacesInsideMarker()
    {
        var marker = ContextIdMarker.Parse("Titolo { # IDH_X = 42 }");

        Assert.Equal("IDH_X", marker.Symbol);
        Assert.Equal(42, marker.ExplicitId);
        Assert.Equal("Titolo", marker.CleanText);
    }

    [Fact]
    public void FindMatchReturnsRangeCoveringWholeMarker()
    {
        var text = "Abc {#IDH_X} def";
        var range = ContextIdMarker.FindMatch(text);

        Assert.NotNull(range);
        Assert.Equal("{#IDH_X}", text.Substring(range!.Value.Start, range.Value.Length));
    }
}

public sealed class SluggerTests
{
    [Fact]
    public void ProducesLowercaseHyphenatedSlug()
    {
        var slugger = new Slugger();
        Assert.Equal("installazione-guidata", slugger.Slug("Installazione Guidata"));
    }

    [Fact]
    public void RemovesDiacritics()
    {
        var slugger = new Slugger();
        Assert.Equal("periferica", slugger.Slug("Perifèrica"));
    }

    [Fact]
    public void EnsuresUniqueness()
    {
        var slugger = new Slugger();
        var first = slugger.Slug("Capitolo");
        var second = slugger.Slug("Capitolo");

        Assert.Equal("capitolo", first);
        Assert.Equal("capitolo-2", second);
    }

    [Fact]
    public void FallsBackToSectionForSymbolOnlyText()
    {
        var slugger = new Slugger();
        Assert.Equal("section", slugger.Slug("###"));
    }
}
