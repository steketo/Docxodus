// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Docxodus.Internal;
using Docxodus.Tests.Ir;
using Xunit;

namespace Docxodus.Tests;

/// <summary>
/// A figure is removed and the next one takes its number: the original's "Figure 24: Legend…" caption goes, and
/// "Figure 25: Details…" becomes "Figure 24: Details…". A cross-reference bookmark both documents carry encloses
/// "Figure 24" (label and <c>SEQ</c> field): in the original on the legend's caption, in the revision on the
/// details' caption. The bookmark must sit on text that BOTH resolutions keep, so that accepting the redline puts
/// it on the details' "Figure 24" and rejecting it on the legend's "Figure 24". Word does this by keeping the
/// legend's "Figure 24" as unchanged text in the deleted caption, whose paragraph mark is deleted, and deleting the
/// label and field of the caption that follows: accepting joins the two paragraphs into "Figure 24: Details…".
/// Kept only on inserted text, the bookmark is lost when Word rejects the redline, and every <c>REF</c> to it prints
/// "Error! Reference source not found.".
/// </summary>
public class DocxDiffRenumberedCaptionBookmarkTests
{
    private const string Legend = ": Legend of the general site plans";
    private const string Details = ": Construction details of the green belts, with planting distances and species: type A on the left, type B on the right";
    private const string Rows = ": Construction details of the tree rows, with planting distances and species: type C";

    private static string R(string text) => $"<w:r><w:t xml:space=\"preserve\">{text}</w:t></w:r>";

    private static string Field(string instruction, string result) =>
        "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>" +
        $"<w:r><w:instrText xml:space=\"preserve\"> {instruction} </w:instrText></w:r>" +
        "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>" + R(result) +
        "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>";

    /// <summary>A figure caption: label and <c>SEQ</c> field, enclosed by the given bookmarks, then the text.</summary>
    private static string Caption(string number, string text, params string[] bookmarks) =>
        "<w:p>" +
        string.Concat(bookmarks.Select((b, i) => $"<w:bookmarkStart w:id=\"{10 + i}\" w:name=\"{b}\"/>")) +
        R("Figure ") + Field("SEQ Figure \\* ARABIC", number) +
        string.Concat(bookmarks.Select((_, i) => $"<w:bookmarkEnd w:id=\"{10 + i}\"/>")) +
        R(text) + "</w:p>";

    private static string Image(string relId, int id) =>
        "<w:p><w:r><w:drawing><wp:inline><wp:extent cx=\"100\" cy=\"200\"/>" +
        $"<wp:docPr id=\"{id}\" name=\"Picture {id}\"/>" +
        "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        $"<pic:nvPicPr><pic:cNvPr id=\"{id}\" name=\"image{id}.png\"/><pic:cNvPicPr/></pic:nvPicPr>" +
        $"<pic:blipFill><a:blip r:embed=\"{relId}\"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
        "<pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"200\"/></a:xfrm>" +
        "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>" +
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";

    private static byte[] Png(byte tag) => IrTestDocuments.TinyPng.Concat(new[] { tag }).ToArray();

    private static string Intro => "<w:p>" + R("The belts follow the site boundary.") + "</w:p>";

    private static string References(params string[] bookmarks) =>
        "<w:p>" + R("As shown in ") +
        string.Join(R(" and "), bookmarks.Select(b => Field($"REF {b} \\h", "Figure 24"))) +
        R(", the belts are wide.") + "</w:p>";

    /// <summary>The shapes. "figures between the captions": each caption follows its image, and the image parts are
    /// renumbered in the revision (one image fewer), so no image paragraph is equal and the diff rewrites the
    /// captions whole, the inserted caption before the deleted ones. "captions only": the details' caption is
    /// paired with its renumbered copy, right after the deleted legend's caption.</summary>
    public static IEnumerable<object[]> Shapes() => new[]
    {
        new object[] { "figures between the captions", new[] { "_Ref_fig_legend" } },
        new object[] { "figures between the captions, two bookmarks", new[] { "_Ref_fig_legend", "_Ref_fig_types_a_b" } },
        new object[] { "captions only", new[] { "_Ref_fig_legend" } },
    };

    private static (WmlDocument Original, WmlDocument Revised) Build(string shape, string[] bookmarks)
    {
        if (shape.StartsWith("figures", StringComparison.Ordinal))
        {
            var original = IrTestDocuments.FromBodyXmlWithImageParts(
                Intro + Image("rIdA", 1) + Caption("24", Legend, bookmarks) + Image("rIdB", 2) +
                Caption("25", Details) + Image("rIdC", 3) + Caption("26", Rows) + References(bookmarks),
                ("rIdA", Png(1)), ("rIdB", Png(2)), ("rIdC", Png(3)));
            var revised = IrTestDocuments.FromBodyXmlWithImageParts(
                Intro + Image("rIdB", 1) + Caption("24", Details, bookmarks) + Image("rIdC", 2) +
                Caption("25", Rows) + References(bookmarks),
                ("rIdB", Png(2)), ("rIdC", Png(3)));
            return (original, revised);
        }
        return (
            IrTestDocuments.FromBodyXml(Intro + Caption("24", Legend, bookmarks) + Caption("25", Details) +
                                        Caption("26", Rows) + References(bookmarks)),
            IrTestDocuments.FromBodyXml(Intro + Caption("24", Details, bookmarks) + Caption("25", Rows) +
                                        References(bookmarks)));
    }

    /// <summary>Accepting the redline gives the revision's paragraphs and puts each bookmark once on the revision's
    /// "Figure 24" of the details' caption; rejecting it gives the original's paragraphs and puts each bookmark once
    /// on the original's "Figure 24" of the legend's caption. The redline itself carries each name once.</summary>
    [Theory]
    [MemberData(nameof(Shapes))]
    public void Compare_BookmarkOnARenumberedCaption_EnclosesTheSameTextInBothViews(string shape, string[] bookmarks)
    {
        var (original, revised) = Build(shape, bookmarks);

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        foreach (var name in bookmarks)
        {
            int copies = Body(redline).Descendants(W.bookmarkStart).Count(s => (string?)s.Attribute(W.name) == name);
            if (copies != 1)
                failures.Add($"redline: {name} on {copies} bookmarkStart(s)");
        }
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), bookmarks);
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), bookmarks);
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, $"[{shape}]\n" + string.Join("\n", failures));
    }

    /// <summary>Several figures follow the removed one, each renumbered (the real report's shape): the moved
    /// bookmark lands on the renumbered caption, and the bookmarks of the later captions, which stay where they are,
    /// keep enclosing their own caption's label and number in both views.</summary>
    [Fact]
    public void Compare_BookmarksOnCaptionsRenumberedInCascade_EncloseTheSameTextInBothViews()
    {
        const string trees = ": Construction details of the hedges along the canal: type D";
        var names = new[] { "_Ref_fig_legend", "_Ref_fig_rows", "_Ref_fig_hedges" };
        var original = IrTestDocuments.FromBodyXmlWithImageParts(
            Intro + Image("rIdA", 1) + Caption("24", Legend, names[0]) + Image("rIdB", 2) + Caption("25", Details) +
            Image("rIdC", 3) + Caption("26", Rows, names[1]) + Image("rIdD", 4) + Caption("27", trees, names[2]) +
            References(names),
            ("rIdA", Png(1)), ("rIdB", Png(2)), ("rIdC", Png(3)), ("rIdD", Png(4)));
        var revised = IrTestDocuments.FromBodyXmlWithImageParts(
            Intro + Image("rIdB", 1) + Caption("24", Details, names[0]) + Image("rIdC", 2) +
            Caption("25", Rows, names[1]) + Image("rIdD", 3) + Caption("26", trees, names[2]) + References(names),
            ("rIdB", Png(2)), ("rIdC", Png(3)), ("rIdD", Png(4)));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), names);
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), names);
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>A second common bookmark opens where the moved one does and ends elsewhere: on the label only
    /// (nested), or on the whole caption (Word's own caption bookmarks). Rewriting the caption for the moved
    /// bookmark would leave the other one around deleted text on accept, so these captions keep the earlier
    /// handling: accepting puts both bookmarks on the revision's text, and both views keep their paragraphs.</summary>
    [Theory]
    [InlineData("label only")]
    [InlineData("whole caption")]
    public void Compare_SecondBookmarkFromTheSameStart_KeepsTheEarlierHandling(string second)
    {
        const string moved = "_Ref_fig_legend";
        const string other = "_Ref_other";
        string Shape(string number, string text) => second == "label only"
            ? $"<w:p><w:bookmarkStart w:id=\"10\" w:name=\"{moved}\"/><w:bookmarkStart w:id=\"11\" w:name=\"{other}\"/>" +
              R("Figure ") + "<w:bookmarkEnd w:id=\"11\"/>" + Field("SEQ Figure \\* ARABIC", number) +
              "<w:bookmarkEnd w:id=\"10\"/>" + R(text) + "</w:p>"
            : $"<w:p><w:bookmarkStart w:id=\"10\" w:name=\"{moved}\"/><w:bookmarkStart w:id=\"11\" w:name=\"{other}\"/>" +
              R("Figure ") + Field("SEQ Figure \\* ARABIC", number) + "<w:bookmarkEnd w:id=\"10\"/>" + R(text) +
              "<w:bookmarkEnd w:id=\"11\"/></w:p>";
        var original = IrTestDocuments.FromBodyXml(Intro + Shape("24", Legend) + Caption("25", Details) +
                                                   Caption("26", Rows) + References(moved, other));
        var revised = IrTestDocuments.FromBodyXml(Intro + Shape("24", Details) + Caption("25", Rows) +
                                                  References(moved, other));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), new[] { moved, other });
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), Array.Empty<string>());
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, $"[{second}]\n" + string.Join("\n", failures));
    }

    /// <summary>Control: the revision makes the label bold. Accepting keeps the original's "Figure 24" when the two
    /// captions are joined, so the join would print the label in the original's formatting; the redline must give
    /// each view its own formatting instead.</summary>
    [Fact]
    public void Compare_BookmarkOnARenumberedCaption_LabelFormattingChanged_EachViewKeepsItsFormatting()
    {
        const string bookmark = "_Ref_fig_legend";
        var boldLabel = R("Figure ").Replace("<w:r>", "<w:r><w:rPr><w:b/></w:rPr>", StringComparison.Ordinal);
        var original = IrTestDocuments.FromBodyXml(Intro + Caption("24", Legend, bookmark) + Caption("25", Details) +
                                                   Caption("26", Rows) + References(bookmark));
        var revised = IrTestDocuments.FromBodyXml(
            Intro + Caption("24", Details, bookmark).Replace(R("Figure "), boldLabel, StringComparison.Ordinal) +
            Caption("25", Rows) + References(bookmark));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        Assert.Equal(string.Join("\n", StyledParagraphs(revised.DocumentByteArray)),
            string.Join("\n", StyledParagraphs(DocxDiffOps.AcceptRevisions(redline))));
        Assert.Equal(string.Join("\n", StyledParagraphs(original.DocumentByteArray)),
            string.Join("\n", StyledParagraphs(DocxDiffOps.RejectRevisions(redline))));
    }

    /// <summary>Control: the legend's caption has its own paragraph properties. Joining it to the next caption on
    /// accept would leave the joined paragraph with one of the two sets, and which one Word keeps is not something
    /// this test can observe, so the captions are not joined: the legend's caption stays deleted whole.</summary>
    [Fact]
    public void Compare_BookmarkOnARenumberedCaption_ParagraphPropertiesDiffer_DoesNotJoinTheCaptions()
    {
        const string bookmark = "_Ref_fig_legend";
        var centred = Caption("24", Legend, bookmark)
            .Replace("<w:p>", "<w:p><w:pPr><w:jc w:val=\"center\"/></w:pPr>", StringComparison.Ordinal);
        var original = IrTestDocuments.FromBodyXml(Intro + centred + Caption("25", Details) + Caption("26", Rows) +
                                                   References(bookmark));
        var revised = IrTestDocuments.FromBodyXml(Intro + Caption("24", Details, bookmark) + Caption("25", Rows) +
                                                  References(bookmark));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var legend = Body(redline).Elements(W.p).Single(p =>
            string.Concat(p.Descendants().Where(e => e.Name == W.t || e.Name == W.delText).Select(e => e.Value))
                .Contains("Legend", StringComparison.Ordinal));
        Assert.Empty(legend.Descendants(W.t));
    }

    private static void CompareView(List<string> failures, string view, byte[] expected, byte[] actual, string[] bookmarks)
    {
        var want = string.Join("\n", Paragraphs(expected));
        var got = string.Join("\n", Paragraphs(actual));
        if (want != got)
            failures.Add($"{view}: paragraphs\nexpected:\n{want}\ngot:\n{got}");
        foreach (var name in bookmarks)
        {
            var (wantCount, wantText, wantHost) = Enclosed(expected, name);
            var (gotCount, gotText, gotHost) = Enclosed(actual, name);
            if (gotCount != wantCount || gotText != wantText || gotHost != wantHost)
                failures.Add($"{view}: {name} expected {wantCount}x «{wantText}» in «{wantHost}», " +
                             $"got {gotCount}x «{gotText}» in «{gotHost}»");
        }
    }

    /// <summary>Revision ids and bookmark ids are unique in the redline, and the redline adds no schema error the
    /// inputs lack.</summary>
    private static void CheckIdsAndSchema(List<string> failures, byte[] redline, WmlDocument original, WmlDocument revised)
    {
        CheckUniqueIds(failures, Body(redline));
        var known = PackageValidation.ErrorKeys(original, FileFormatVersions.Office2019, PackageValidation.Keys.IdAtXPathDescription);
        known.UnionWith(PackageValidation.ErrorKeys(revised, FileFormatVersions.Office2019, PackageValidation.Keys.IdAtXPathDescription));
        var fresh = PackageValidation.ErrorKeys(redline, FileFormatVersions.Office2019, PackageValidation.Keys.IdAtXPathDescription)
            .Where(e => !known.Contains(e)).ToList();
        if (fresh.Count > 0)
            failures.Add("schema: " + string.Join(" | ", fresh.Take(5)));
    }

    private static void CheckUniqueIds(List<string> failures, XElement body)
    {
        foreach (var kind in new[] { W.ins, W.del, W.bookmarkStart, W.bookmarkEnd })
        {
            var repeated = body.Descendants(kind).Select(e => (string?)e.Attribute(W.id)).GroupBy(i => i)
                .Where(g => g.Count() > 1).Select(g => g.Key).ToList();
            if (repeated.Count > 0)
                failures.Add($"{kind.LocalName}: repeated w:id [{string.Join(",", repeated)}]");
        }
    }

    // ---- the bookmark pass on hand-written redlines ----------------------------------------------

    private int _revisionId = 100;

    private string Mark(string kind) =>
        $"<w:pPr><w:rPr><w:{kind} w:id=\"{_revisionId++}\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\"/></w:rPr></w:pPr>";

    private string Tracked(string kind, string text) =>
        $"<w:{kind} w:id=\"{_revisionId++}\" w:author=\"A\" w:date=\"2026-01-01T00:00:00Z\">" +
        (kind == "del" ? $"<w:r><w:delText xml:space=\"preserve\">{text}</w:delText></w:r>" : R(text)) + $"</w:{kind}>";

    /// <summary>A paragraph inserted or deleted whole: the label, enclosed by the bookmark when one is named, then
    /// the rest.</summary>
    private string Whole(string kind, string? bookmark, int bookmarkId, string label, string rest) =>
        "<w:p>" + Mark(kind) +
        (bookmark == null ? "" : $"<w:bookmarkStart w:id=\"{bookmarkId}\" w:name=\"{bookmark}\"/>") +
        Tracked(kind, label) + (bookmark == null ? "" : $"<w:bookmarkEnd w:id=\"{bookmarkId}\"/>") +
        Tracked(kind, rest) + "</w:p>";

    /// <summary>Run the pass on a hand-written redline body, then compare each view's paragraphs with the expected
    /// ones, the named bookmarks with their expected paragraph in each view (around some text), and the ids.</summary>
    private static void AssertPassGives(string body, string[] names, string[] original, string[] revised,
        (string Name, string Accepted, string Rejected)[] enclosed)
    {
        var document = IrTestDocuments.FromBodyXml(body);
        byte[] bytes;
        using (var stream = new MemoryStream())
        {
            stream.Write(document.DocumentByteArray, 0, document.DocumentByteArray.Length);
            using (var package = WordprocessingDocument.Open(stream, true))
            {
                var main = package.MainDocumentPart!;
                var root = main.GetXDocument().Root!.Element(W.body)!;
                int next = 1000;
                Docxodus.Ir.Diff.IrMarkupRenderer.CarryBookmarksOntoCommonText(
                    root.Descendants(W.bookmarkStart).ToList(), root.Descendants(W.bookmarkEnd).ToList(),
                    new HashSet<string>(names), new HashSet<string>(names),
                    Docxodus.Ir.Diff.IrMarkupRenderer.CitedBookmarkNames(new[] { root }), () => next++);
                main.PutXDocument();
            }
            bytes = stream.ToArray();
        }

        var failures = new List<string>();
        var accepted = DocxDiffOps.AcceptRevisions(bytes);
        var rejected = DocxDiffOps.RejectRevisions(bytes);
        if (string.Join("\n", Paragraphs(accepted)) != string.Join("\n", revised))
            failures.Add("accept:\n" + string.Join("\n", Paragraphs(accepted)));
        if (string.Join("\n", Paragraphs(rejected)) != string.Join("\n", original))
            failures.Add("reject:\n" + string.Join("\n", Paragraphs(rejected)));
        foreach (var (name, acc, rej) in enclosed)
        {
            var (_, accText, accHost) = Enclosed(accepted, name);
            var (_, rejText, rejHost) = Enclosed(rejected, name);
            if (accHost != acc || rejHost != rej || accText.Length == 0 || rejText.Length == 0)
                failures.Add($"{name}: accept «{accText}» in «{accHost}», reject «{rejText}» in «{rejHost}»");
        }
        CheckUniqueIds(failures, Body(bytes));
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>Two bookmarks on equal labels, their paragraphs crossed: inserted A, deleted B, deleted A, inserted
    /// B. Rewriting A moves both deleted paragraphs before inserted A and makes A's label unchanged, so B's deleted
    /// paragraph is no longer followed only by whole-deleted ones: B must be checked again on the rewritten markup,
    /// not on the markup the candidates were found in.</summary>
    [Fact]
    public void Pass_CrossedCandidates_InsertedFirst_EachViewKeepsItsParagraphs() =>
        AssertPassGives(
            Whole("ins", "X", 2, "Figure 7", ": New A") + Whole("del", "Y", 3, "Figure 7", ": Old B") +
            Whole("del", "X", 1, "Figure 7", ": Old A") + Whole("ins", "Y", 4, "Figure 7", ": New B"),
            new[] { "X", "Y" },
            original: new[] { "Figure 7: Old B", "Figure 7: Old A" },
            revised: new[] { "Figure 7: New A", "Figure 7: New B" },
            enclosed: new[] { ("X", "Figure 7: New A", "Figure 7: Old A") });

    /// <summary>Deleted A, deleted B, inserted B, inserted A: rewriting A moves inserted B before deleted A, so for B
    /// the inserted paragraph now comes first, past a deleted paragraph that is no longer whole-deleted.</summary>
    [Fact]
    public void Pass_CrossedCandidates_DeletedFirst_EachViewKeepsItsParagraphs() =>
        AssertPassGives(
            Whole("del", "X", 1, "Figure 7", ": Old A") + Whole("del", "Y", 3, "Figure 7", ": Old B") +
            Whole("ins", "Y", 4, "Figure 7", ": New B") + Whole("ins", "X", 2, "Figure 7", ": New A"),
            new[] { "X", "Y" },
            original: new[] { "Figure 7: Old A", "Figure 7: Old B" },
            revised: new[] { "Figure 7: New B", "Figure 7: New A" },
            enclosed: new[] { ("X", "Figure 7: New A", "Figure 7: Old A") });

    /// <summary>The deleted paragraph comes first and a whole-inserted paragraph stands between it and the revised
    /// one: the inserted paragraph moves before the deleted one, so accepting joins the label to the revised
    /// paragraph and not to the inserted one.</summary>
    [Fact]
    public void Pass_DeletedFirstWithAnInsertedParagraphBetween_MovesTheInsertedParagraphOut() =>
        AssertPassGives(
            Whole("del", "X", 1, "Figure 7", ": Old A") + Whole("ins", null, 0, "Inserted", " note") +
            Whole("ins", "X", 2, "Figure 7", ": New A"),
            new[] { "X" },
            original: new[] { "Figure 7: Old A" },
            revised: new[] { "Inserted note", "Figure 7: New A" },
            enclosed: new[] { ("X", "Figure 7: New A", "Figure 7: Old A") });

    /// <summary>Run the pass on a hand-written redline body and check it leaves the markup as it was.</summary>
    private static void AssertPassLeavesAlone(string body, string[] leftNames, string[] rightNames)
    {
        var root = XElement.Parse($"<w:body xmlns:w=\"{W.w.NamespaceName}\">{body}</w:body>");
        var before = root.ToString(SaveOptions.DisableFormatting);
        int next = 1000;
        bool changed = Docxodus.Ir.Diff.IrMarkupRenderer.CarryBookmarksOntoCommonText(
            root.Descendants(W.bookmarkStart).ToList(), root.Descendants(W.bookmarkEnd).ToList(),
            new HashSet<string>(leftNames), new HashSet<string>(rightNames),
            Docxodus.Ir.Diff.IrMarkupRenderer.CitedBookmarkNames(new[] { root }), () => next++);
        Assert.False(changed);
        Assert.Equal(before, root.ToString(SaveOptions.DisableFormatting));
    }

    /// <summary>The whole-inserted paragraph between the two holds the start of a bookmark that ends in a later
    /// paragraph: moving it would change what that bookmark spans, so the pair is left alone.</summary>
    [Fact]
    public void Pass_InsertedParagraphBetweenOpensARangeEndingElsewhere_LeavesThePairAlone()
    {
        var inserted = "<w:p>" + Mark("ins") + "<w:bookmarkStart w:id=\"9\" w:name=\"Z\"/>" +
                       Tracked("ins", "Inserted note") + "</w:p>";
        AssertPassLeavesAlone(
            Whole("del", "X", 1, "Figure 7", ": Old A") + inserted + Whole("ins", "X", 2, "Figure 7", ": New A") +
            "<w:p><w:r><w:t>Tail</w:t></w:r><w:bookmarkEnd w:id=\"9\"/></w:p>",
            new[] { "X" }, new[] { "X", "Z" });
    }

    /// <summary>A bookmark only the original has spans the deleted caption, label included: making the label
    /// unchanged would carry that bookmark's text into the accepted document, so the pair is left alone.</summary>
    [Fact]
    public void Pass_OriginalOnlyBookmarkOverTheDeletedLabel_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("<w:bookmarkStart w:id=\"1\"", "<w:bookmarkStart w:id=\"8\" w:name=\"Q\"/><w:bookmarkStart w:id=\"1\"", StringComparison.Ordinal)
            .Replace("</w:p>", "<w:bookmarkEnd w:id=\"8\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(deleted + Whole("ins", "X", 2, "Figure 7", ": New A"), new[] { "X", "Q" }, new[] { "X" });
    }

    /// <summary>A bookmark only the revision has opens with the revised label and spans the whole revised caption:
    /// deleting that label would leave the bookmark without it on accept, so the pair is left alone.</summary>
    [Fact]
    public void Pass_RevisionOnlyBookmarkOverTheRevisedCaption_LeavesThePairAlone()
    {
        var inserted = Whole("ins", "X", 2, "Figure 7", ": New A")
            .Replace("<w:bookmarkStart w:id=\"2\"", "<w:bookmarkStart w:id=\"7\" w:name=\"N\"/><w:bookmarkStart w:id=\"2\"", StringComparison.Ordinal)
            .Replace("</w:p>", "<w:bookmarkEnd w:id=\"7\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(Whole("del", "X", 1, "Figure 7", ": Old A") + inserted, new[] { "X" }, new[] { "X", "N" });
    }

    /// <summary>Word's hidden bookmarks (a <c>_Hlk</c> left by a copy, the <c>_GoBack</c> of the last edit) sit in
    /// the captions outside the label no field cites: they do not keep the caption out of the rewrite, and both views
    /// still give their document's paragraphs with the moved bookmark on its label.</summary>
    [Theory]
    [InlineData("_Hlk in the original's legend text")]
    [InlineData("_GoBack before the revised label")]
    public void Compare_HiddenBookmarkOutsideTheLabel_StillMovesTheBookmark(string place)
    {
        const string bookmark = "_Ref_fig_legend";
        var legend = Caption("24", Legend, bookmark);
        var details = Caption("24", Details, bookmark);
        if (place.StartsWith("_Hlk", StringComparison.Ordinal))
            legend = legend.Replace(R(Legend), R(": Legend of the ") + "<w:bookmarkStart w:id=\"40\" w:name=\"_Hlk212345678\"/>" +
                                    R("general") + "<w:bookmarkEnd w:id=\"40\"/>" + R(" site plans"), StringComparison.Ordinal);
        else
            details = details.Replace("<w:p>", "<w:p><w:bookmarkStart w:id=\"40\" w:name=\"_GoBack\"/><w:bookmarkEnd w:id=\"40\"/>",
                                      StringComparison.Ordinal);
        var original = IrTestDocuments.FromBodyXml(Intro + legend + Caption("25", Details) + Caption("26", Rows) + References(bookmark));
        var revised = IrTestDocuments.FromBodyXml(Intro + details + Caption("25", Rows) + References(bookmark));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), new[] { bookmark });
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), new[] { bookmark });
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, $"[{place}]\n" + string.Join("\n", failures));
    }

    /// <summary>A hidden-looking bookmark that a field cites is a cross-reference target like any other: it keeps the
    /// pair out of the rewrite.</summary>
    [Fact]
    public void Pass_CitedUnderscoreBookmarkInTheDeletedCaption_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("</w:p>", "<w:bookmarkStart w:id=\"8\" w:name=\"_Hlk1\"/><w:bookmarkEnd w:id=\"8\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(
            deleted + Whole("ins", "X", 2, "Figure 7", ": New A") +
            "<w:p>" + Field("REF _Hlk1 \\h", "Figure 7") + "</w:p>",
            new[] { "X", "_Hlk1" }, new[] { "X" });
    }

    /// <summary>An uncited hidden bookmark inside the deleted label would be carried into the accepted document with
    /// the label: it keeps the pair out like any other bookmark.</summary>
    [Fact]
    public void Pass_HiddenBookmarkInsideTheDeletedLabel_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("<w:bookmarkEnd w:id=\"1\"/>", "<w:bookmarkStart w:id=\"8\" w:name=\"_Hlk2\"/><w:bookmarkEnd w:id=\"8\"/><w:bookmarkEnd w:id=\"1\"/>",
                     StringComparison.Ordinal);
        AssertPassLeavesAlone(deleted + Whole("ins", "X", 2, "Figure 7", ": New A"), new[] { "X", "_Hlk2" }, new[] { "X" });
    }

    /// <summary>A cross-reference bookmark of Word (<c>_Ref…</c>) opens before the moved one and spans the whole
    /// caption, and only a footnote cites it: it is a target like any other, so the caption keeps the earlier
    /// handling, and accepting leaves both bookmarks on the revision's text.</summary>
    [Fact]
    public void Compare_RefBookmarkOnTheWholeCaptionCitedOnlyByAFootnote_KeepsTheEarlierHandling()
    {
        const string moved = "_Ref_fig_legend";
        const string whole = "_Ref123456";
        static string Shape(string number, string text) =>
            $"<w:p><w:bookmarkStart w:id=\"11\" w:name=\"{whole}\"/><w:bookmarkStart w:id=\"10\" w:name=\"{moved}\"/>" +
            R("Figure ") + Field("SEQ Figure \\* ARABIC", number) + "<w:bookmarkEnd w:id=\"10\"/>" + R(text) +
            "<w:bookmarkEnd w:id=\"11\"/></w:p>";
        var note = "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
        var footnote = R("Compare ") + Field($"REF {whole} \\h", "Figure 24") + R(".");
        var original = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            Intro + Shape("24", Legend) + Caption("25", Details) + Caption("26", Rows) + References(moved) + note, footnote);
        var revised = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            Intro + Shape("24", Details) + Caption("25", Rows) + References(moved) + note, footnote);

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), new[] { moved, whole });
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), Array.Empty<string>());
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>An uncited <c>_Hlk</c> opens before the deleted label and closes after it: its range spans the copy,
    /// so making the label unchanged would carry it into the accepted document, and the pair is left alone.</summary>
    [Fact]
    public void Pass_HiddenBookmarkSpanningTheDeletedLabel_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("<w:bookmarkStart w:id=\"1\"", "<w:bookmarkStart w:id=\"8\" w:name=\"_Hlk3\"/><w:bookmarkStart w:id=\"1\"", StringComparison.Ordinal)
            .Replace("</w:p>", "<w:bookmarkEnd w:id=\"8\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(deleted + Whole("ins", "X", 2, "Figure 7", ": New A"), new[] { "X", "_Hlk3" }, new[] { "X" });
    }

    /// <summary>A field holding only a bookmark's name (<c>{ _Hlk4 }</c>, Word's implicit REF) cites it: the hidden
    /// bookmark after the deleted label keeps the pair out.</summary>
    [Fact]
    public void Pass_HiddenBookmarkCitedByABareNameField_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("</w:p>", "<w:bookmarkStart w:id=\"8\" w:name=\"_Hlk4\"/><w:bookmarkEnd w:id=\"8\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(
            deleted + Whole("ins", "X", 2, "Figure 7", ": New A") +
            "<w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> _Hlk4 </w:instrText></w:r>" +
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>Old A</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>",
            new[] { "X", "_Hlk4" }, new[] { "X" });
    }

    /// <summary>Only <c>_GoBack</c> and <c>_Hlk…</c> are let through: a <c>_Ref…</c> after the deleted label keeps
    /// the pair out even when no field the pass reads cites it.</summary>
    [Fact]
    public void Pass_UncitedRefBookmarkAfterTheDeletedLabel_LeavesThePairAlone()
    {
        var deleted = Whole("del", "X", 1, "Figure 7", ": Old A")
            .Replace("</w:p>", "<w:bookmarkStart w:id=\"8\" w:name=\"_Ref9\"/><w:bookmarkEnd w:id=\"8\"/></w:p>", StringComparison.Ordinal);
        AssertPassLeavesAlone(deleted + Whole("ins", "X", 2, "Figure 7", ": New A"), new[] { "X", "_Ref9" }, new[] { "X" });
    }

    /// <summary>An <c>_Hlk</c> in the legend's text that only a footnote cites: the citing fields are read in every
    /// story, so it keeps the caption out, and the legend's label stays deleted in the redline.</summary>
    [Fact]
    public void Compare_HiddenBookmarkCitedOnlyByAFootnote_KeepsTheEarlierHandling()
    {
        const string bookmark = "_Ref_fig_legend";
        var legend = Caption("24", Legend, bookmark).Replace(R(Legend), R(": Legend of the ") +
            "<w:bookmarkStart w:id=\"40\" w:name=\"_Hlk212345678\"/>" + R("general") + "<w:bookmarkEnd w:id=\"40\"/>" +
            R(" site plans"), StringComparison.Ordinal);
        var note = "<w:p><w:r><w:t xml:space=\"preserve\">See </w:t></w:r><w:r><w:footnoteReference w:id=\"1\"/></w:r></w:p>";
        var footnote = R("Compare ") + Field("REF _Hlk212345678 \\h", "general") + R(".");
        var original = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            Intro + legend + Caption("25", Details) + Caption("26", Rows) + References(bookmark) + note, footnote);
        var revised = IrTestDocuments.FromBodyXmlWithFootnoteParagraph(
            Intro + Caption("24", Details, bookmark) + Caption("25", Rows) + References(bookmark) + note, footnote);

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var legendInRedline = Body(redline).Elements(W.p).Single(p =>
            string.Concat(p.Descendants().Where(e => e.Name == W.t || e.Name == W.delText).Select(e => e.Value))
                .Contains("Legend", StringComparison.Ordinal));
        Assert.Empty(legendInRedline.Descendants(W.t));
        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), new[] { bookmark });
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), Array.Empty<string>());
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>Captions only, renumbered in cascade, and the original's next caption carries its own bookmark on its
    /// label, which the revision keeps on the renumbered caption beside the moved one. That bookmark sits in the
    /// revised caption before the moved copy's end, so the pair keeps the earlier handling (14.1.0.2): accepting is
    /// right for every name and rejecting gives back the original's paragraphs, but on reject both bookmarks are
    /// left around no text, which Word drops. Recorded here so that a change of this outcome is seen.</summary>
    [Fact]
    public void Compare_CaptionsOnlyInCascade_NextCaptionWithItsOwnBookmark_KeepsTheEarlierHandling()
    {
        var names = new[] { "_Ref_fig_legend", "_Ref_fig_details" };
        var original = IrTestDocuments.FromBodyXml(Intro + Caption("24", Legend, names[0]) + Caption("25", Details, names[1]) +
                                                   Caption("26", Rows) + References(names));
        var revised = IrTestDocuments.FromBodyXml(Intro + Caption("24", Details, names) + Caption("25", Rows) + References(names));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), names);
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), Array.Empty<string>());
        foreach (var name in names)
            if (Enclosed(DocxDiffOps.RejectRevisions(redline), name).Text.Length != 0)
                failures.Add($"reject: {name} encloses text: the outcome changed, update this test");
        CheckIdsAndSchema(failures, redline, original, revised);
        Assert.True(failures.Count == 0, string.Join("\n", failures));
    }

    /// <summary>The whole-inserted paragraph between the two closes a range opened earlier and opens another with the
    /// same id (before step (B) the two sources' ids can coincide): its starts and ends do not pair inside it, so it
    /// is not moved and the pair is left alone.</summary>
    [Fact]
    public void Pass_InsertedParagraphBetweenClosesOneRangeAndOpensAnotherWithTheSameId_LeavesThePairAlone()
    {
        var inserted = "<w:p>" + Mark("ins") + "<w:bookmarkEnd w:id=\"5\"/>" + Tracked("ins", "Inserted note") +
                       "<w:bookmarkStart w:id=\"5\" w:name=\"Z2\"/></w:p>";
        AssertPassLeavesAlone(
            "<w:p><w:bookmarkStart w:id=\"5\" w:name=\"Z1\"/><w:r><w:t>Head</w:t></w:r></w:p>" +
            Whole("del", "X", 1, "Figure 7", ": Old A") + inserted + Whole("ins", "X", 2, "Figure 7", ": New A") +
            "<w:p><w:r><w:t>Tail</w:t></w:r><w:bookmarkEnd w:id=\"5\"/></w:p>",
            new[] { "X", "Z1" }, new[] { "X", "Z2" });
    }

    // ---- observers -------------------------------------------------------------------------------

    private static XElement Body(byte[] bytes)
    {
        using var stream = new MemoryStream(bytes);
        using var document = WordprocessingDocument.Open(stream, false);
        return document.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
    }

    /// <summary>Each body paragraph as its text, field instructions in brackets and images as <c>&lt;image&gt;</c>.</summary>
    private static List<string> Paragraphs(byte[] bytes) =>
        Body(bytes).Elements(W.p).Select(p => string.Concat(p.Descendants().Select(e =>
            e.Name == W.t ? e.Value
            : e.Name == W.instrText ? "[" + e.Value.Trim() + "]"
            : e.Name == W.drawing ? "<image>"
            : ""))).ToList();

    /// <summary>Each body paragraph as its text, bold stretches between asterisks.</summary>
    private static List<string> StyledParagraphs(byte[] bytes) =>
        Body(bytes).Elements(W.p).Select(p =>
        {
            var parts = new List<(bool Bold, string Text)>();
            foreach (var t in p.Descendants(W.t))
            {
                bool bold = t.Parent?.Element(W.rPr)?.Element(W.b) != null;
                if (parts.Count > 0 && parts[^1].Bold == bold)
                    parts[^1] = (bold, parts[^1].Text + t.Value);
                else
                    parts.Add((bold, t.Value));
            }
            return string.Concat(parts.Select(x => x.Bold ? "*" + x.Text + "*" : x.Text));
        }).ToList();

    /// <summary>How many bookmarkStarts carry the name; the text between the first one and its end; the text of the
    /// paragraph it opens in. Word drops a bookmark whose text a resolution removes, so a bookmark left around no
    /// text, or around other text, is a bookmark the view has lost.</summary>
    private static (int Count, string Text, string Host) Enclosed(byte[] bytes, string name)
    {
        var body = Body(bytes);
        var starts = body.Descendants(W.bookmarkStart).Where(s => (string?)s.Attribute(W.name) == name).ToList();
        if (starts.Count == 0)
            return (0, "", "");
        var start = starts[0];
        var id = (string?)start.Attribute(W.id);
        var all = body.Descendants().ToList();
        int from = all.IndexOf(start);
        int to = all.FindIndex(from, e => e.Name == W.bookmarkEnd && (string?)e.Attribute(W.id) == id);
        var text = to < 0
            ? "<no end>"
            : string.Concat(all.Skip(from).Take(to - from).Where(e => e.Name == W.t).Select(e => e.Value));
        var host = start.Ancestors(W.p).FirstOrDefault() is { } p
            ? string.Concat(p.Descendants(W.t).Select(t => t.Value))
            : "<between blocks>";
        return (starts.Count, text, host);
    }
}
