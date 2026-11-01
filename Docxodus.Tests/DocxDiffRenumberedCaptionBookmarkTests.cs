// Copyright (c) John Scrudato IV. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
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
        $"<wp:docPr id=\"{id}\" name=\"Picture {id}\"/><a:graphic><a:graphicData>" +
        "<pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        $"<pic:blipFill><a:blip r:embed=\"{relId}\"/></pic:blipFill></pic:pic>" +
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
        Assert.True(failures.Count == 0, $"[{shape}]\n" + string.Join("\n", failures));
    }

    /// <summary>A second bookmark nested in the first, on the label only: the caption is rewritten once, for the
    /// outer bookmark, and the nested one keeps the ordinary handling; both views keep their paragraphs.</summary>
    [Fact]
    public void Compare_NestedBookmarkOnARenumberedCaption_RewritesTheCaptionOnce()
    {
        const string outer = "_Ref_fig_legend";
        static string Nested(string number, string text) =>
            $"<w:p><w:bookmarkStart w:id=\"10\" w:name=\"{outer}\"/><w:bookmarkStart w:id=\"11\" w:name=\"_Ref_label\"/>" +
            R("Figure ") + "<w:bookmarkEnd w:id=\"11\"/>" + Field("SEQ Figure \\* ARABIC", number) +
            "<w:bookmarkEnd w:id=\"10\"/>" + R(text) + "</w:p>";
        var original = IrTestDocuments.FromBodyXml(Intro + Nested("24", Legend) + Caption("25", Details) +
                                                   Caption("26", Rows) + References(outer));
        var revised = IrTestDocuments.FromBodyXml(Intro + Nested("24", Details) + Caption("25", Rows) +
                                                  References(outer));

        var redline = DocxCompare.Compare(original, revised).DocumentByteArray;

        var failures = new List<string>();
        CompareView(failures, "accept", revised.DocumentByteArray, DocxDiffOps.AcceptRevisions(redline), new[] { outer });
        CompareView(failures, "reject", original.DocumentByteArray, DocxDiffOps.RejectRevisions(redline), new[] { outer });
        Assert.True(failures.Count == 0, string.Join("\n", failures));
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
