using Word2Chm.Core.Common;

namespace Word2Chm.Core.Tests;

public sealed class ContextIdMarkerTests
{
    [Fact]
    public void ParsesSymbolOnly()
    {
        var marker = ContextIdMarker.Parse("Installazione {#IDH_INSTALLAZIONE}");

        Assert.Equal("IDH_INSTALLAZIONE", marker.Symbol);
        Assert.Equal("Installazione", marker.CleanText);
    }

    [Fact]
    public void IgnoresALegacyInlineNumber()
    {
        // The number used to be trusted; now the header is the only source, so an old
        // {#IDH_RIFERIMENTI=5000} must still be recognised (and removed) but its value is gone.
        var marker = ContextIdMarker.Parse("Riferimenti {#IDH_RIFERIMENTI=5000}");

        Assert.Equal("IDH_RIFERIMENTI", marker.Symbol);
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

public sealed class ContextIdHeaderTests
{
    [Fact]
    public void ReadsDecimalAndHexadecimalDefines()
    {
        var header = ContextIdHeader.Parse("""
            #define IDH_A 1000
            #define IDH_B 0x0FDA
            """);

        Assert.Equal(1000, header.Definitions["IDH_A"]);
        Assert.Equal(0x0FDA, header.Definitions["IDH_B"]);
    }

    [Fact]
    public void IgnoresTheIncludeGuardAndNonNumericDefines()
    {
        var header = ContextIdHeader.Parse("""
            #ifndef _HELPID_H_
            #define _HELPID_H_           // no number: the guard
            #define IDH_A 1000
            #define IDH_B IDH_A          // expands to another macro
            #endif
            """);

        Assert.True(header.TryGetId("IDH_A", out var id));
        Assert.Equal(1000, id);
        Assert.False(header.TryGetId("_HELPID_H_", out _));
        Assert.False(header.TryGetId("IDH_B", out _));
    }

    [Fact]
    public void IgnoresDefinesInsideLineComments()
    {
        // A commented-out example must not become a real definition, or the ID would win.
        var header = ContextIdHeader.Parse("""
            // #define IDH_A 1
            #define IDH_A 2000
            """);

        Assert.Equal(2000, header.Definitions["IDH_A"]);
    }

    [Fact]
    public void AcceptsSuffixesAndParentheses()
    {
        var header = ContextIdHeader.Parse("""
            #define IDH_A 1000u
            #define IDH_B (2000)
            #define IDH_C 3000L
            """);

        Assert.Equal(1000, header.Definitions["IDH_A"]);
        Assert.Equal(2000, header.Definitions["IDH_B"]);
        Assert.Equal(3000, header.Definitions["IDH_C"]);
    }

    [Fact]
    public void LastDefinitionWins()
    {
        // C allows redefining a macro and the last one is effective; the header must agree.
        var header = ContextIdHeader.Parse("""
            #define IDH_A 1000
            #define IDH_A 2000
            """);

        Assert.Equal(2000, header.Definitions["IDH_A"]);
    }

    [Fact]
    public void MissingSymbolIsReportedNotGuessed()
    {
        var header = ContextIdHeader.Parse("#define IDH_A 1000");

        Assert.False(header.TryGetId("IDH_MANCANTE", out _));
    }

    [Fact]
    public void ReadsEnumeratorsWithExplicitValues()
    {
        var header = ContextIdHeader.Parse("""
            enum class HelpID {
                IDH_EDIT_PARAMETERS = 1000,
                IDH_START_JOB = 1001,
                IDH_AXES = 0x03EA,
            };
            """);

        Assert.Equal(1000, header.Definitions["IDH_EDIT_PARAMETERS"]);
        Assert.Equal(1001, header.Definitions["IDH_START_JOB"]);
        Assert.Equal(1002, header.Definitions["IDH_AXES"]);
    }

    [Fact]
    public void ImplicitEnumeratorsIncrementThePreviousValue()
    {
        var header = ContextIdHeader.Parse("""
            enum class HelpID {
                IDH_EDIT_PARAMETERS = 1000,
                IDH_START_JOB,
                IDH_AXES,
                IDH_BATCH,
                IDH_SCRIPT,
                IDH_SPINDLE_GROUP,
                IDH_ATC,
            };
            """);

        Assert.Equal(1000, header.Definitions["IDH_EDIT_PARAMETERS"]);
        Assert.Equal(1001, header.Definitions["IDH_START_JOB"]);
        Assert.Equal(1002, header.Definitions["IDH_AXES"]);
        Assert.Equal(1003, header.Definitions["IDH_BATCH"]);
        Assert.Equal(1004, header.Definitions["IDH_SCRIPT"]);
        Assert.Equal(1005, header.Definitions["IDH_SPINDLE_GROUP"]);
        Assert.Equal(1006, header.Definitions["IDH_ATC"]);
    }

    [Fact]
    public void AnEnumeratorWithoutValueAfterAnExpressionKeepsTheRun()
    {
        var header = ContextIdHeader.Parse("""
            #define IDH_BASE 1000
            enum Plain {
                IDH_A = IDH_BASE + 1,
                IDH_B,
                IDH_C = 2000,
                IDH_D,
            };
            """);

        Assert.Equal(1001, header.Definitions["IDH_A"]);
        Assert.Equal(1002, header.Definitions["IDH_B"]);
        Assert.Equal(2000, header.Definitions["IDH_C"]);
        Assert.Equal(2001, header.Definitions["IDH_D"]);
    }

    [Fact]
    public void ResolvesCommonConstantExpressions()
    {
        var header = ContextIdHeader.Parse("""
            #define IDH_BASE 0x1000
            enum Flags {
                IDH_ONE = 1 << 4,
                IDH_TWO = IDH_BASE | 0x2,
                IDH_THREE = (2 + 3) * 4,
                IDH_FOUR = 0b1010,
                IDH_FIVE = 010,
            };
            """);

        Assert.Equal(16, header.Definitions["IDH_ONE"]);
        Assert.Equal(0x1002, header.Definitions["IDH_TWO"]);
        Assert.Equal(20, header.Definitions["IDH_THREE"]);
        Assert.Equal(10, header.Definitions["IDH_FOUR"]);
        Assert.Equal(8, header.Definitions["IDH_FIVE"]);
    }

    [Fact]
    public void SkipsEnumeratorsItCannotResolveInsteadOfGuessing()
    {
        var header = ContextIdHeader.Parse("""
            enum Weird {
                IDH_A = someFunction(),
                IDH_B,
                IDH_C = 3000,
                IDH_D,
            };
            """);

        // The unresolvable value poisons the implicit run after it, but a later explicit value
        // starts a fresh one.
        Assert.False(header.TryGetId("IDH_A", out _));
        Assert.False(header.TryGetId("IDH_B", out _));
        Assert.Equal(3000, header.Definitions["IDH_C"]);
        Assert.Equal(3001, header.Definitions["IDH_D"]);
    }

    [Fact]
    public void MixesDefinesAndEnumsKeepingTheLastValue()
    {
        var header = ContextIdHeader.Parse("""
            #define IDH_A 1000
            enum Plain {
                IDH_A = 2000,
                IDH_B,
            };
            #define IDH_B 3000
            """);

        Assert.Equal(2000, header.Definitions["IDH_A"]);
        Assert.Equal(3000, header.Definitions["IDH_B"]);
    }

    [Fact]
    public void IgnoresBlockCommentsSpanningLines()
    {
        var header = ContextIdHeader.Parse("""
            /* enum Plain {
                   IDH_A = 1000,
               }; */
            enum Plain {
                IDH_A = 2000,
                IDH_B,
            };
            """);

        Assert.Equal(2000, header.Definitions["IDH_A"]);
        Assert.Equal(2001, header.Definitions["IDH_B"]);
    }

    [Fact]
    public void ReadsAnEnumInsideANamespace()
    {
        // Wrapping the enum in a namespace (or a class) is common; the symbol names are the
        // same, so the namespace around them changes nothing.
        var header = ContextIdHeader.Parse("""
            #pragma once

            namespace HelpID
            {
                enum {
                    IDH_EDIT_PARAMETERS = 1000,
                    IDH_START_JOB,
                    IDH_AXES,
                    IDH_BATCH,
                    IDH_SCRIPT,
                    IDH_SPINDLE_GROUP,
                    IDH_ATC,
                };
            }
            """);

        Assert.Equal(1000, header.Definitions["IDH_EDIT_PARAMETERS"]);
        Assert.Equal(1003, header.Definitions["IDH_BATCH"]);
        Assert.Equal(1006, header.Definitions["IDH_ATC"]);
    }

    [Fact]
    public void ReadsAWholeEnumWrittenOnOneLine()
    {
        var header = ContextIdHeader.Parse("namespace HelpID { enum { IDH_A = 1000, IDH_B, IDH_C }; }");

        Assert.Equal(1000, header.Definitions["IDH_A"]);
        Assert.Equal(1001, header.Definitions["IDH_B"]);
        Assert.Equal(1002, header.Definitions["IDH_C"]);
    }

    [Fact]
    public void ReadsSeveralEnumsInTheSameScope()
    {
        var header = ContextIdHeader.Parse("""
            namespace HelpID {
                enum { IDH_A = 100, IDH_B, };
                enum Other { IDH_X = 200, IDH_Y, };
            }
            """);

        Assert.Equal(101, header.Definitions["IDH_B"]);
        Assert.Equal(201, header.Definitions["IDH_Y"]);
    }

    [Fact]
    public void ThrowsWhenAnEnumValueExceedsTheIntRange()
    {
        var text = """
            enum TooBig {
                IDH_A = 2147483648,
            };
            """;

        Assert.Throws<OverflowException>(() => ContextIdHeader.Parse(text));
    }

    [Fact]
    public void ReadsTheSampleHeaderShippedWithTheProject()
    {
        // Guards the contract with the real header used by the sample document.
        var path = Path.Combine(FindRepositoryRoot(), "sample", "Output", "SkipperQt_IT.h");
        var header = ContextIdHeader.Load(path);

        Assert.Equal(1000, header.Definitions["IDH_EDIT_PARAMETERS"]);
        Assert.Equal(1001, header.Definitions["IDH_START_JOB"]);
        Assert.Equal(1002, header.Definitions["IDH_AXES"]);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Word2Chm.sln")))
        {
            directory = directory.Parent;
        }

        Assert.NotNull(directory);
        return directory!.FullName;
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
