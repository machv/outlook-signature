/***************************************************************************

Copyright (c) Microsoft Corporation 2013.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

***************************************************************************/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public class FormattingAssemblerSettings
    {
        public bool RemoveStyleNamesFromParagraphAndRunProperties;
        public bool ClearStyles;
        public bool OrderElementsPerStandard;
        public bool CreatePtFontNameAttribute;
        public bool RestrictToSupportedNumberingFormats;
        public bool RestrictToSupportedLanguages;

        public FormattingAssemblerSettings()
        {
            RemoveStyleNamesFromParagraphAndRunProperties = true;
            ClearStyles = true;
            OrderElementsPerStandard = true;
            CreatePtFontNameAttribute = true;
            RestrictToSupportedNumberingFormats = true;
            RestrictToSupportedLanguages = true;
        }
    }

    public static class FormattingAssembler
    {
        public static WmlDocument AssembleFormatting(WmlDocument document, FormattingAssemblerSettings settings)
        {
            using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(document))
            {
                using (WordprocessingDocument doc = streamDoc.GetWordprocessingDocument())
                {
                    AssembleFormatting(doc, settings);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void AssembleFormatting(WordprocessingDocument wDoc, FormattingAssemblerSettings settings)
        {
            FormattingAssemblerInfo fai = new FormattingAssemblerInfo();
            XDocument sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            XElement defaultParagraphStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "paragraph");
            if (defaultParagraphStyle != null)
                fai.DefaultParagraphStyleName = (string)defaultParagraphStyle.Attribute(W.styleId);
            XElement defaultCharacterStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "character");
            if (defaultCharacterStyle != null)
                fai.DefaultCharacterStyleName = (string)defaultCharacterStyle.Attribute(W.styleId);
            XElement defaultTableStyle = sXDoc
                .Root
                .Elements(W.style)
                .FirstOrDefault(st => st.Attribute(W._default).ToBoolean() == true &&
                    (string)st.Attribute(W.type) == "table");
            if (defaultTableStyle != null)
                fai.DefaultTableStyleName = (string)defaultTableStyle.Attribute(W.styleId);
            AssembleListItemInformation(wDoc);
            foreach (var part in wDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                FixNonconformantHexValues(pxd.Root);
                AnnotateWithGlobalDefaults(wDoc, pxd.Root);
                AnnotateTablesWithTableStyles(wDoc, pxd.Root);
                AnnotateParagraphs(fai, wDoc, pxd.Root, settings);
                AnnotateRuns(fai, wDoc, pxd.Root, settings);
            }
            NormalizeListItems(fai, wDoc, settings);
            if (settings.ClearStyles)
                ClearStyles(wDoc);
            foreach (var part in wDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                pxd.Root.Descendants().Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
                FormattingAssembler.NormalizePropsForPart(pxd, settings);
                var newRoot = (XElement)CleanupTransform(pxd.Root);
                pxd.Root.ReplaceWith(newRoot);
                part.PutXDocument();
            }
        }

        private static void FixNonconformantHexValues(XElement root)
        {
            foreach (var tblLook in root.Descendants(W.tblLook))
            {
                if (tblLook.Attributes().Any(a => a.Name != W.val))
                    continue;
                if (tblLook.Attribute(W.val) == null)
                    continue;
                string hexValue = tblLook.Attribute(W.val).Value;
                int val = int.Parse(hexValue, System.Globalization.NumberStyles.HexNumber);
                tblLook.Add(new XAttribute(W.firstRow, (val & 0x0020) != 0 ? "1" : "0"));
                tblLook.Add(new XAttribute(W.lastRow, (val & 0x0040) != 0 ? "1" : "0"));
                tblLook.Add(new XAttribute(W.firstColumn, (val & 0x0080) != 0 ? "1" : "0"));
                tblLook.Add(new XAttribute(W.lastColumn, (val & 0x0100) != 0 ? "1" : "0"));
                tblLook.Add(new XAttribute(W.noHBand, (val & 0x0200) != 0 ? "1" : "0"));
                tblLook.Add(new XAttribute(W.noVBand, (val & 0x0400) != 0 ? "1" : "0"));
            }
            foreach (var cnfStyle in root.Descendants(W.cnfStyle))
            {
                if (cnfStyle.Attributes().Any(a => a.Name != W.val))
                    continue;
                if (cnfStyle.Attribute(W.val) == null)
                    continue;
                var va = cnfStyle.Attribute(W.val).Value.ToArray();
                cnfStyle.Add(new XAttribute(W.firstRow, va[0]));
                cnfStyle.Add(new XAttribute(W.lastRow, va[1]));
                cnfStyle.Add(new XAttribute(W.firstColumn, va[2]));
                cnfStyle.Add(new XAttribute(W.lastColumn, va[3]));
                cnfStyle.Add(new XAttribute(W.oddVBand, va[4]));
                cnfStyle.Add(new XAttribute(W.evenVBand, va[5]));
                cnfStyle.Add(new XAttribute(W.oddHBand, va[6]));
                cnfStyle.Add(new XAttribute(W.evenHBand, va[7]));
                cnfStyle.Add(new XAttribute(W.firstRowLastColumn, va[8]));
                cnfStyle.Add(new XAttribute(W.firstRowFirstColumn, va[9]));
                cnfStyle.Add(new XAttribute(W.lastRowLastColumn, va[10]));
                cnfStyle.Add(new XAttribute(W.lastRowFirstColumn, va[11]));
            }
        }

        private static object CleanupTransform(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.tabs && element.Element(W.tab) == null)
                    return null;

                if (element.Name == W.tblStyleRowBandSize || element.Name == W.tblStyleColBandSize)
                    return null;

                // a cleaner solution would be to not include the w:ins and w:del elements when rolling up the paragraph run properties into
                // the run properties.
                if ((element.Name == W.ins || element.Name == W.del) && element.Parent.Name == W.rPr)
                {
                    if (element.Parent.Parent.Name == W.r || element.Parent.Parent.Name == W.rPrChange)
                        return null;
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CleanupTransform(n)));
            }
            return node;
        }

        private static void ClearStyles(WordprocessingDocument wDoc)
        {
            var stylePart = wDoc.MainDocumentPart.StyleDefinitionsPart;
            var sXDoc = stylePart.GetXDocument();

            var newRoot = new XElement(sXDoc.Root.Name,
                sXDoc.Root.Attributes(),
                sXDoc.Root.Elements().Select(e =>
                {
                    if (e.Name != W.style)
                        return e;
                    return new XElement(e.Name,
                        e.Attributes(),
                        e.Element(W.name),
                        new XElement(W.pPr),
                        new XElement(W.rPr));
                }));

            var globalrPr = newRoot
                .Elements(W.docDefaults)
                .Elements(W.rPrDefault)
                .Elements(W.rPr)
                .FirstOrDefault();
            if (globalrPr != null)
                globalrPr.ReplaceWith(new XElement(W.rPr));

            var globalpPr = newRoot
                .Elements(W.docDefaults)
                .Elements(W.pPrDefault)
                .Elements(W.pPr)
                .FirstOrDefault();
            if (globalpPr != null)
                globalpPr.ReplaceWith(new XElement(W.pPr));

            sXDoc.Root.ReplaceWith(newRoot);

            stylePart.PutXDocument();
        }

        private static void NormalizeListItems(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, FormattingAssemblerSettings settings)
        {
            foreach (var part in wDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                XElement newRoot = (XElement)NormalizeListItemsTransform(fai, wDoc, pxd.Root, settings);
                pxd.Root.ReplaceWith(newRoot);
            }
        }

        private static object NormalizeListItemsTransform(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XNode node, FormattingAssemblerSettings settings)
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.p)
                {
                    var li = ListItemRetriever.RetrieveListItem(wDoc, element, null);
                    if (li != null)
                    {
                        ListItemRetriever.ListItemInfo listItemInfo = element.Annotation<ListItemRetriever.ListItemInfo>();

                        var newParaProps = new XElement(W.pPr,
                            element.Attributes(),
                            element.Elements(W.pPr).Elements().Where(e => e.Name != W.numPr)
                        );

                        XElement listItemRunProps = null;
                        if (listItemInfo != null)
                        {
                            var paraStyleRunProps = CharStyleRollup(fai, wDoc, element);

                            var lvlStyleName = (string)listItemInfo
                                .Lvl
                                .Elements(W.pStyle)
                                .Attributes(W.val)
                                .FirstOrDefault();

                            if (lvlStyleName == null)
                                lvlStyleName = (string)wDoc
                                    .MainDocumentPart
                                    .StyleDefinitionsPart
                                    .GetXDocument()
                                    .Root
                                    .Elements(W.style)
                                    .Where(s => (string)s.Attribute(W.type) == "paragraph" && s.Attribute(W._default).ToBoolean() == true)
                                    .Attributes(W.styleId)
                                    .FirstOrDefault();

                            XElement lvlStyleRpr = ParaStyleRunPropsStack(wDoc, lvlStyleName)
                                .Aggregate(new XElement(W.rPr),
                                    (r, s) =>
                                    {
                                        var newCharStyleRunProps = MergeStyleElement(s, r);
                                        return newCharStyleRunProps;
                                    });

                            var mergedRunProps = MergeStyleElement(lvlStyleRpr, paraStyleRunProps);

                            // pick up everything except bold & italic from accumulated run props
                            var accumulatedRunProps = element.Elements(PtOpenXml.pt + "pPr").Elements(W.rPr).FirstOrDefault();
                            if (accumulatedRunProps != null)
                            {
                                accumulatedRunProps = new XElement(W.rPr, accumulatedRunProps
                                    .Elements()
                                    .Where(e => e.Name != W.b && e.Name != W.bCs && e.Name != W.i && e.Name != W.szCs));
                                mergedRunProps = MergeStyleElement(accumulatedRunProps, mergedRunProps);
                            }
                            var listItemLvlRunProps = listItemInfo.Lvl.Elements(W.rPr).FirstOrDefault();
                            listItemRunProps = MergeStyleElement(listItemLvlRunProps, mergedRunProps);
                        }

                        var listItemRun = new XElement(W.r,
                            listItemRunProps,
                            new XElement(W.t,
                                new XAttribute(XNamespace.Xml + "space", "preserve"),
                                li));
                        AdjustFontAttributes(wDoc, listItemRun, listItemRunProps, settings);
                        XElement newPara = new XElement(W.p,
                            newParaProps,
                            listItemRun,
                            new XElement(W.r,
                                new XElement(W.tab)),
                            element.Elements().Where(e => e.Name != W.pPr).Select(n => NormalizeListItemsTransform(fai, wDoc, n, settings)));
                        return newPara;

                    }
                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => NormalizeListItemsTransform(fai, wDoc, n, settings)));
            }
            return node;
        }

        public static void NormalizePropsForPart(XDocument pxd, FormattingAssemblerSettings settings)
        {
            if (settings.CreatePtFontNameAttribute)
            {
                pxd.Root.Descendants().Attributes().Where(d => d.Name.Namespace == PtOpenXml.pt && d.Name.LocalName != "FontName").Remove();
                if (pxd.Root.Attribute(XNamespace.Xmlns + "pt") == null)
                    pxd.Root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.pt.NamespaceName));
                if (pxd.Root.Attribute(XNamespace.Xmlns + "mc") == null)
                    pxd.Root.Add(new XAttribute(XNamespace.Xmlns + "mc", MC.mc.NamespaceName));
                XAttribute mci = pxd.Root.Attribute(MC.Ignorable);
                if (mci != null)
                {
                    var ig = pxd.Root.Attribute(MC.Ignorable).Value + " pt14";
                    mci.Value = ig;
                }
                else
                {
                    pxd.Root.Add(new XAttribute(MC.Ignorable, "pt14"));
                }
            }
            else
            {
                pxd.Root.Descendants().Attributes().Where(d => d.Name.Namespace == PtOpenXml.pt).Remove();
            }
            var runProps = pxd.Root.Descendants(PtOpenXml.pt + "rPr").ToList();
            foreach (var item in runProps)
            {
                XElement newRunProps = new XElement(W.rPr,
                    item.Attributes(),
                    item.Elements());
                XElement parent = item.Parent;
                if (parent.Name == W.p)
                {
                    XElement existingParaProps = parent.Element(W.pPr);
                    if (existingParaProps == null)
                    {
                        existingParaProps = new XElement(W.pPr);
                        parent.Add(existingParaProps);
                    }
                    XElement existingRunProps = existingParaProps.Element(W.rPr);
                    if (existingRunProps != null)
                    {
                        if (!settings.RemoveStyleNamesFromParagraphAndRunProperties)
                        {
                            if (newRunProps.Element(W.rStyle) == null)
                                newRunProps.Add(existingRunProps.Element(W.rStyle));
                        }
                        existingRunProps.ReplaceWith(newRunProps);
                    }
                    else
                        existingParaProps.Add(newRunProps);
                }
                else
                {
                    XElement existingRunProps = parent.Element(W.rPr);
                    if (existingRunProps != null)
                    {
                        if (!settings.RemoveStyleNamesFromParagraphAndRunProperties)
                        {
                            if (newRunProps.Element(W.rStyle) == null)
                                newRunProps.Add(existingRunProps.Element(W.rStyle));
                        }
                        existingRunProps.ReplaceWith(newRunProps);
                    }
                    else
                        parent.Add(newRunProps);
                }
            }
            var paraProps = pxd.Root.Descendants(PtOpenXml.pt + "pPr").ToList();
            foreach (var item in paraProps)
            {
                var paraRunProps = item.Parent.Elements(W.pPr).Elements(W.rPr).FirstOrDefault();
                var merged = MergeStyleElement(item.Element(W.rPr), paraRunProps);
                if (!settings.RemoveStyleNamesFromParagraphAndRunProperties)
                {
                    if (merged.Element(W.rStyle) == null)
                    {
                        merged.Add(paraRunProps.Element(W.rStyle));
                    }
                }

                XElement newParaProps = new XElement(W.pPr,
                    item.Attributes(),
                    item.Elements().Where(e => e.Name != W.rPr),
                    merged);
                XElement para = item.Parent;
                XElement existingParaProps = para.Element(W.pPr);
                if (existingParaProps != null)
                {
                    if (!settings.RemoveStyleNamesFromParagraphAndRunProperties)
                    {
                        if (newParaProps.Element(W.pStyle) == null)
                            newParaProps.Add(existingParaProps.Element(W.pStyle));
                    }
                    existingParaProps.ReplaceWith(newParaProps);
                }
                else
                    para.Add(newParaProps);
            }
            var tblProps = pxd.Root.Descendants(PtOpenXml.pt + "tblPr").ToList();
            foreach (var item in tblProps)
            {
                XElement newTblProps = new XElement(item);
                newTblProps.Name = W.tblPr;
                XElement table = item.Parent;
                XElement existingTableProps = table.Element(W.tblPr);
                if (existingTableProps != null)
                    existingTableProps.ReplaceWith(newTblProps);
                else
                    table.AddFirst(newTblProps);
            }
            var trProps = pxd.Root.Descendants(PtOpenXml.pt + "trPr").ToList();
            foreach (var item in trProps)
            {
                XElement newTrProps = new XElement(item);
                newTrProps.Name = W.trPr;
                XElement row = item.Parent;
                XElement existingRowProps = row.Element(W.trPr);
                if (existingRowProps != null)
                    existingRowProps.ReplaceWith(newTrProps);
                else
                    row.AddFirst(newTrProps);
            }
            var tcProps = pxd.Root.Descendants(PtOpenXml.pt + "tcPr").ToList();
            foreach (var item in tcProps)
            {
                XElement newTcProps = new XElement(item);
                newTcProps.Name = W.tcPr;
                XElement row = item.Parent;
                XElement existingRowProps = row.Element(W.tcPr);
                if (existingRowProps != null)
                    existingRowProps.ReplaceWith(newTcProps);
                else
                    row.AddFirst(newTcProps);
            }
            pxd.Root.Descendants(W.numPr).Remove();
            if (settings.RemoveStyleNamesFromParagraphAndRunProperties)
            {
                pxd.Root.Descendants(W.pStyle).Where(ps => ps.Parent.Name == W.pPr).Remove();
                pxd.Root.Descendants(W.rStyle).Where(ps => ps.Parent.Name == W.rPr).Remove();
            }
            pxd.Root.Descendants(W.tblStyle).Where(ps => ps.Parent.Name == W.tblPr).Remove();
            pxd.Root.Descendants().Where(d => d.Name.Namespace == PtOpenXml.pt).Remove();
            if (settings.OrderElementsPerStandard)
            {
                XElement newRoot = (XElement)TransformAndOrderElements(pxd.Root);
                pxd.Root.ReplaceWith(newRoot);
            }
        }

        private static Dictionary<XName, int> Order_pPr = new Dictionary<XName, int>
        {
            { W.pStyle, 10 },
            { W.keepNext, 20 },
            { W.keepLines, 30 },
            { W.pageBreakBefore, 40 },
            { W.framePr, 50 },
            { W.widowControl, 60 },
            { W.numPr, 70 },
            { W.suppressLineNumbers, 80 },
            { W.pBdr, 90 },
            { W.shd, 100 },
            { W.tabs, 120 },
            { W.suppressAutoHyphens, 130 },
            { W.kinsoku, 140 },
            { W.wordWrap, 150 },
            { W.overflowPunct, 160 },
            { W.topLinePunct, 170 },
            { W.autoSpaceDE, 180 },
            { W.autoSpaceDN, 190 },
            { W.bidi, 200 },
            { W.adjustRightInd, 210 },
            { W.snapToGrid, 220 },
            { W.spacing, 230 },
            { W.ind, 240 },
            { W.contextualSpacing, 250 },
            { W.mirrorIndents, 260 },
            { W.suppressOverlap, 270 },
            { W.jc, 280 },
            { W.textDirection, 290 },
            { W.textAlignment, 300 },
            { W.textboxTightWrap, 310 },
            { W.outlineLvl, 320 },
            { W.divId, 330 },
            { W.cnfStyle, 340 },
            { W.rPr, 350 },
            { W.sectPr, 360 },
            { W.pPrChange, 370 },
        };

        private static Dictionary<XName, int> Order_rPr = new Dictionary<XName, int>
        {
            { W.ins, 10 },
            { W.del, 20 },
            { W.rStyle, 30 },
            { W.rFonts, 40 },
            { W.b, 50 },
            { W.bCs, 60 },
            { W.i, 70 },
            { W.iCs, 80 },
            { W.caps, 90 },
            { W.smallCaps, 100 },
            { W.strike, 110 },
            { W.dstrike, 120 },
            { W.outline, 130 },
            { W.shadow, 140 },
            { W.emboss, 150 },
            { W.imprint, 160 },
            { W.noProof, 170 },
            { W.snapToGrid, 180 },
            { W.vanish, 190 },
            { W.webHidden, 200 },
            { W.color, 210 },
            { W.spacing, 220 },
            { W._w, 230 },
            { W.kern, 240 },
            { W.position, 250 },
            { W.sz, 260 },
            { W14.wShadow, 270 },
            { W14.wTextOutline, 280 },
            { W14.wTextFill, 290 },
            { W14.wScene3d, 300 },
            { W14.wProps3d, 310 },
            { W.szCs, 320 },
            { W.highlight, 330 },
            { W.u, 340 },
            { W.effect, 350 },
            { W.bdr, 360 },
            { W.shd, 370 },
            { W.fitText, 380 },
            { W.vertAlign, 390 },
            { W.rtl, 400 },
            { W.cs, 410 },
            { W.em, 420 },
            { W.lang, 430 },
            { W.eastAsianLayout, 440 },
            { W.specVanish, 450 },
            { W.oMath, 460 },
        };

        private static Dictionary<XName, int> Order_tblPr = new Dictionary<XName, int>
        {
            { W.tblStyle, 10 },
            { W.tblpPr, 20 },
            { W.tblOverlap, 30 },
            { W.bidiVisual, 40 },
            { W.tblStyleRowBandSize, 50 },
            { W.tblStyleColBandSize, 60 },
            { W.tblW, 70 },
            { W.jc, 80 },
            { W.tblCellSpacing, 90 },
            { W.tblInd, 100 },
            { W.tblBorders, 110 },
            { W.shd, 120 },
            { W.tblLayout, 130 },
            { W.tblCellMar, 140 },
            { W.tblLook, 150 },
            { W.tblCaption, 160 },
            { W.tblDescription, 170 },
        };

        private static Dictionary<XName, int> Order_tblBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.start, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
        };

        private static Dictionary<XName, int> Order_tcPr = new Dictionary<XName, int>
        {
            { W.cnfStyle, 10 },
            { W.tcW, 20 },
            { W.gridSpan, 30 },
            { W.hMerge, 40 },
            { W.vMerge, 50 },
            { W.tcBorders, 60 },
            { W.shd, 70 },
            { W.noWrap, 80 },
            { W.tcMar, 90 },
            { W.textDirection, 100 },
            { W.tcFitText, 110 },
            { W.vAlign, 120 },
            { W.hideMark, 130 },
            { W.headers, 140 },
        };

        private static Dictionary<XName, int> Order_tcBorders = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.start, 20 },
            { W.left, 30 },
            { W.bottom, 40 },
            { W.right, 50 },
            { W.end, 60 },
            { W.insideH, 70 },
            { W.insideV, 80 },
            { W.tl2br, 90 },
            { W.tr2bl, 100 },
        };

        private static Dictionary<XName, int> Order_pBdr = new Dictionary<XName, int>
        {
            { W.top, 10 },
            { W.left, 20 },
            { W.bottom, 30 },
            { W.right, 40 },
            { W.between, 50 },
            { W.bar, 60 },
        };

        private static object TransformAndOrderElements(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                if (element.Name == W.pPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e => {
                            if (Order_pPr.ContainsKey(e.Name))
                                return Order_pPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.rPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_rPr.ContainsKey(e.Name))
                                return Order_rPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tblPr.ContainsKey(e.Name))
                                return Order_tblPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcPr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tcPr.ContainsKey(e.Name))
                                return Order_tcPr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tcBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tcBorders.ContainsKey(e.Name))
                                return Order_tcBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.tblBorders)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_tblBorders.ContainsKey(e.Name))
                                return Order_tblBorders[e.Name];
                            return 999;
                        }));

                if (element.Name == W.pBdr)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements().Select(e => (XElement)TransformAndOrderElements(e)).OrderBy(e =>
                        {
                            if (Order_pBdr.ContainsKey(e.Name))
                                return Order_pBdr[e.Name];
                            return 999;
                        }));

                if (element.Name == W.p)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.pPr).Select(e => (XElement)TransformAndOrderElements(e)),
                        element.Elements().Where(e => e.Name != W.pPr).Select(e => (XElement)TransformAndOrderElements(e)));

                if (element.Name == W.r)
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Elements(W.rPr).Select(e => (XElement)TransformAndOrderElements(e)),
                        element.Elements().Where(e => e.Name != W.rPr).Select(e => (XElement)TransformAndOrderElements(e)));

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformAndOrderElements(n)));
            }
            return node;
        }

        private static void Simplify(XElement xElement)
        {
            xElement.Descendants().Attributes(W14.paraId).Remove();
            xElement.Descendants().Attributes(W14.textId).Remove();
            xElement.Descendants().Attributes(W.rsidR).Remove();
            xElement.Descendants().Attributes(W.rsidRDefault).Remove();
            xElement.Descendants().Attributes(W.rsidP).Remove();
            xElement.Descendants().Attributes(W.rsidRPr).Remove();
        }

        private static void AssembleListItemInformation(WordprocessingDocument wordDoc)
        {
            foreach (var part in wordDoc.ContentParts())
            {
                XDocument xDoc = part.GetXDocument();
                foreach (var para in xDoc.Descendants(W.p))
                {
                    ListItemRetriever.RetrieveListItem(wordDoc, para, "");
                }
            }
        }

        private static void AnnotateWithGlobalDefaults(WordprocessingDocument wDoc, XElement rootElement)
        {
            XElement globalDefaultParaProps = null;
            XElement globalDefaultParaPropsAsDefined = null;
            XElement globalDefaultRunProps = null;
            XElement globalDefaultRunPropsAsDefined = null;
            XDocument sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            XElement docDefaults = sXDoc.Root.Element(W.docDefaults);
            if (docDefaults != null)
            {
                globalDefaultParaPropsAsDefined = docDefaults.Elements(W.pPrDefault).Elements(W.pPr)
                    .FirstOrDefault();
                if (globalDefaultParaPropsAsDefined == null)
                    globalDefaultParaPropsAsDefined = new XElement(W.pPr,
                        new XElement(W.rPr));
                globalDefaultRunPropsAsDefined = docDefaults.Elements(W.rPrDefault).Elements(W.rPr)
                    .FirstOrDefault();
                if (globalDefaultRunPropsAsDefined == null)
                    globalDefaultRunPropsAsDefined = new XElement(W.rPr);
                var runPropsForGlobalDefaultParaProps = MergeStyleElement(globalDefaultRunPropsAsDefined, globalDefaultParaPropsAsDefined.Element(W.rPr));
                globalDefaultParaProps = new XElement(globalDefaultParaPropsAsDefined.Name,
                    globalDefaultParaPropsAsDefined.Attributes(),
                    globalDefaultParaPropsAsDefined.Elements().Where(e => e.Name != W.rPr),
                    runPropsForGlobalDefaultParaProps);
                globalDefaultRunProps = MergeStyleElement(globalDefaultParaPropsAsDefined.Element(W.rPr), globalDefaultRunPropsAsDefined);
            }
            if (globalDefaultParaProps == null)
            {
                globalDefaultParaProps = new XElement(W.pPr);
            }
            if (globalDefaultRunProps == null)
            {
                globalDefaultRunProps = new XElement(W.rPr);
            }
            XElement ptGlobalDefaultParaProps = new XElement(globalDefaultParaProps);
            XElement ptGlobalDefaultRunProps = new XElement(globalDefaultRunProps);
            ptGlobalDefaultParaProps.Name = PtOpenXml.pt + "pPr";
            ptGlobalDefaultRunProps.Name = PtOpenXml.pt + "rPr";
            var parasAndRuns = rootElement.Descendants().Where(d =>
            {
                return d.Name == W.p || d.Name == W.r;
            });
            foreach (var d in parasAndRuns)
            {
                if (d.Name == W.p)
                {
                    d.Add(ptGlobalDefaultParaProps);
                }
                else
                {
                    d.Add(ptGlobalDefaultRunProps);
                }
            }
        }

        private static void AnnotateTablesWithTableStyles(WordprocessingDocument wDoc, XElement rootElement)
        {
            XDocument sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            foreach (var tbl in rootElement.Descendants(W.tbl))
            {
                string tblStyleName = (string)tbl.Elements(W.tblPr).Elements(W.tblStyle).Attributes(W.val).FirstOrDefault();
                if (tblStyleName != null)
                {
                    XElement style = TableStyleRollup(wDoc, tblStyleName);

                    // annotate table with table style, in PowerTools namespace
                    style.Name = PtOpenXml.pt + "style";
                    tbl.Add(style);

                    XElement tblPr2 = style.Element(W.tblPr);
                    XElement tblPr3 = MergeStyleElement(tbl.Element(W.tblPr), tblPr2);

                    if (tblPr3 != null)
                    {
                        XElement newTblPr = new XElement(tblPr3);
                        newTblPr.Name = PtOpenXml.pt + "tblPr";
                        tbl.Add(newTblPr);
                    }

                    // Iterate through every row and cell in the table, rolling up row properties and cell properties
                    // as appropriate per the cnfStyle element, then replacing the row and cell properties
                    foreach (var row in tbl.Elements(W.tr))
                    {
                        XElement trPr2 = null;
                        trPr2 = style.Element(W.trPr);
                        if (trPr2 == null)
                            trPr2 = new XElement(W.trPr);
                        XElement rowCnf = row.Elements(W.trPr).Elements(W.cnfStyle).FirstOrDefault();
                        if (rowCnf != null)
                        {
                            foreach (var ot in TableStyleOverrideTypes)
                            {
                                XName attName = TableStyleOverrideXNameMap[ot];
                                if (attName == null ||
                                    (rowCnf != null && rowCnf.Attribute(attName).ToBoolean() == true))
                                {
                                    XElement o = style
                                        .Elements(W.tblStylePr)
                                        .Where(tsp => (string)tsp.Attribute(W.type) == ot)
                                        .FirstOrDefault();
                                    if (o != null)
                                    {
                                        XElement ottrPr = o.Element(W.trPr);
                                        trPr2 = MergeStyleElement(ottrPr, trPr2);
                                    }
                                }
                            }
                        }
                        trPr2 = MergeStyleElement(row.Element(W.trPr), trPr2);
                        if (trPr2.HasElements)
                        {
                            trPr2.Name = PtOpenXml.pt + "trPr";
                            row.Add(trPr2);
                        }
                    }
                    foreach (var cell in tbl.Elements(W.tr).Elements(W.tc))
                    {
                        XElement tcPr2 = null;
                        tcPr2 = style.Element(W.tcPr);
                        if (tcPr2 == null)
                            tcPr2 = new XElement(W.tcPr);
                        XElement rowCnf = cell.Ancestors(W.tr).Take(1).Elements(W.trPr).Elements(W.cnfStyle).FirstOrDefault();
                        XElement cellCnf = cell.Elements(W.tcPr).Elements(W.cnfStyle).FirstOrDefault();
                        foreach (var ot in TableStyleOverrideTypes)
                        {
                            XName attName = TableStyleOverrideXNameMap[ot];
                            if (attName == null ||
                                (rowCnf != null && rowCnf.Attribute(attName).ToBoolean() == true) ||
                                (cellCnf != null && cellCnf.Attribute(attName).ToBoolean() == true))
                            {
                                XElement o = style
                                    .Elements(W.tblStylePr)
                                    .Where(tsp => (string)tsp.Attribute(W.type) == ot)
                                    .FirstOrDefault();
                                if (o != null)
                                {
                                    XElement ottcPr = o.Element(W.tcPr);
                                    tcPr2 = MergeStyleElement(ottcPr, tcPr2);
                                }
                            }
                        }
                        var localTcPr = cell.Element(W.tcPr);
                        tcPr2 = MergeStyleElement(localTcPr, tcPr2);
                        if (tcPr2.HasElements)
                        {
                            tcPr2.Name = PtOpenXml.pt + "tcPr";
                            cell.Add(tcPr2);
                        }
                    }
                }
            }
        }

        private static XElement TableStyleRollup(WordprocessingDocument wDoc, string tblStyleName)
        {
            var tblStyleChain = TableStyleStack(wDoc, tblStyleName)
                .Reverse();
            XElement rolledStyle = new XElement(W.style);
            foreach (var style in tblStyleChain)
            {
                rolledStyle = MergeStyleElement(style, rolledStyle);
            }
            return rolledStyle;
        }

        private static XName[] SpecialCaseChildProperties =
        {
            W.tblPr,
            W.trPr,
            W.tcPr,
            W.pPr,
            W.rPr,
            W.pBdr,
            W.tabs,
            W.rFonts,
            W.ind,
            W.spacing,
            W.tblStylePr,
            W.tcBorders,
            W.tblBorders,
            W.lang,
        };

        private static XName[] MergeChildProperties =
        {
            W.tblPr,
            W.trPr,
            W.tcPr,
            W.pPr,
            W.rPr,
            W.pBdr,
            W.tcBorders,
            W.tblBorders,
        };

        private static string[] TableStyleOverrideTypes =
        {
            "wholeTable",
            "band1Vert",
            "band2Vert",
            "band1Horz",
            "band2Horz",
            "firstCol",
            "lastCol",
            "firstRow",
            "lastRow",
            "neCell",
            "nwCell",
            "seCell",
            "swCell",
        };

        private static Dictionary<string, XName> TableStyleOverrideXNameMap = new Dictionary<string,XName>
        {
            {"wholeTable", null},
            {"band1Vert", W.w + "oddVBand"},
            {"band2Vert", W.w + "evenVBand"},
            {"band1Horz", W.w + "oddHBand"},
            {"band2Horz", W.w + "evenHBand"},
            {"firstCol", W.w + "firstColumn"},
            {"lastCol", W.w + "lastColumn"},
            {"firstRow", W.w + "firstRow"},
            {"lastRow", W.w + "lastRow"},
            {"neCell", W.w + "firstRowLastColumn"},
            {"nwCell", W.w + "firstRowFirstColumn"},
            {"seCell", W.w + "lastRowLastColumn"},
            {"swCell", W.w + "lastRowFirstColumn"},
        };

        private static XElement MergeStyleElement(XElement higherPriorityElement, XElement lowerPriorityElement)
        {
            // If, when in the process of merging, the source element doesn't have a
            // corresponding element in the merged element, then include the source element
            // in the merged element.
            if (lowerPriorityElement == null)
                return higherPriorityElement;
            if (higherPriorityElement == null)
                return lowerPriorityElement;

            var hpe = higherPriorityElement
                .Elements()
                .Where(e => !SpecialCaseChildProperties.Contains(e.Name))
                .ToArray();
            var lpe = lowerPriorityElement
                .Elements()
                .Where(e => !SpecialCaseChildProperties.Contains(e.Name) && !hpe.Select(z => z.Name).Contains(e.Name))
                .ToArray();
            var ma = SpacingMerge(higherPriorityElement.Element(W.spacing), lowerPriorityElement.Element(W.spacing));
            var rFonts = FontMerge(higherPriorityElement.Element(W.rFonts), lowerPriorityElement.Element(W.rFonts));
            var tabs = TabsMerge(higherPriorityElement.Element(W.tabs), lowerPriorityElement.Element(W.tabs));
            var ind = IndMerge(higherPriorityElement.Element(W.ind), lowerPriorityElement.Element(W.ind));
            var lang = LangMerge(higherPriorityElement.Element(W.lang), lowerPriorityElement.Element(W.lang));
            var mcp = MergeChildProperties
                .Select(e =>
                {
                    // test is here to prevent unnecessary recursion to make debugging easier
                    var h = higherPriorityElement.Element(e);
                    var l = lowerPriorityElement.Element(e);
                    if (h == null && l == null)
                        return null;
                    if (h == null && l != null)
                        return l;
                    if (h != null && l == null)
                        return h;
                    return MergeStyleElement(h, l);
                })
                .Where(m => m != null)
                .ToArray();
            var tsor = TableStyleOverrideTypes
                .Select(e =>
                {
                    // test is here to prevent unnecessary recursion to make debugging easier
                    var h = higherPriorityElement.Elements(W.tblStylePr).FirstOrDefault(tsp => (string)tsp.Attribute(W.type) == e);
                    var l = lowerPriorityElement.Elements(W.tblStylePr).FirstOrDefault(tsp => (string)tsp.Attribute(W.type) == e);
                    if (h == null && l == null)
                        return null;
                    if (h == null && l != null)
                        return l;
                    if (h != null && l == null)
                        return h;
                    return MergeStyleElement(h, l);
                })
                .Where(m => m != null)
                .ToArray();

            XElement newMergedElement = new XElement(higherPriorityElement.Name,
                new XAttribute(XNamespace.Xmlns + "w", W.w),
                higherPriorityElement.Attributes().Where(a => !a.IsNamespaceDeclaration),
                hpe,  // higher priority elements
                lpe,  // lower priority elements where there is not a higher priority element of same name
                ind,  // w:ind has very special rules
                ma,   // elements that require merged attributes
                lang,
                rFonts,  // font merge is special case
                tabs,    // tabs merge is special case
                mcp,  // elements that need child properties to be merged
                tsor // merged table style override elements
            );

            return newMergedElement;
        }

        private static XElement LangMerge(XElement hLang, XElement lLang)
        {
            if (hLang == null && lLang == null)
                return null;
            if (hLang != null && lLang == null)
                return hLang;
            if (lLang != null && hLang == null)
                return lLang;
            return new XElement(W.lang,
                hLang.Attribute(W.val) != null ? hLang.Attribute(W.val) : lLang.Attribute(W.val),
                hLang.Attribute(W.bidi) != null ? hLang.Attribute(W.bidi) : lLang.Attribute(W.bidi),
                hLang.Attribute(W.eastAsia) != null ? hLang.Attribute(W.eastAsia) : lLang.Attribute(W.eastAsia));
        }

        private enum IndAttType
        {
            End,
            FirstLineOrHanging,
            Start,
            Left,
            Right,
            None,
        };

        private class IndAttInfo
        {
            public XAttribute Attribute;
            public IndAttType AttributeType;
            public int HighPri;
            public int LowPri;
        }

        private static XElement IndMerge(XElement higherPriorityElement, XElement lowerPriorityElement)
        {
            if (higherPriorityElement == null && lowerPriorityElement == null)
                return null;
            if (higherPriorityElement != null && lowerPriorityElement == null)
                return higherPriorityElement;
            if (lowerPriorityElement != null && higherPriorityElement == null)
                return lowerPriorityElement;
            var hpa = higherPriorityElement.Attributes().Select(a => GetIndAttInfo(a, 1)).ToList();
            var lpa = lowerPriorityElement.Attributes().Select(a => GetIndAttInfo(a, 2)).ToList();
            var atts = hpa.Concat(lpa)
                .GroupBy(iai => iai.AttributeType)
                .Select(g => g.OrderBy(c => c.HighPri).ThenBy(c => c.LowPri).First().Attribute).ToList();
            var newInd = new XElement(W.ind, atts);
            return newInd;
        }

        private static IndAttInfo GetIndAttInfo(XAttribute a, int highPri)
        {
            IndAttType iat = IndAttType.None;
            int lowPri = 0;
            if (a.Name == W.right)
            {
                iat = IndAttType.Right;
                lowPri = 2;
            }
            if (a.Name == W.rightChars)
            {
                iat = IndAttType.Right;
                lowPri = 1;
            }
            if (a.Name == W.left)
            {
                iat = IndAttType.Left;
                lowPri = 2;
            }
            if (a.Name == W.leftChars)
            {
                iat = IndAttType.Left;
                lowPri = 1;
            }
            if (a.Name == W.start)
            {
                iat = IndAttType.Start;
                lowPri = 2;
            }
            if (a.Name == W.startChars)
            {
                iat = IndAttType.Start;
                lowPri = 1;
            }
            if (a.Name == W.end)
            {
                iat = IndAttType.End;
                lowPri = 2;
            }
            if (a.Name == W.endChars)
            {
                iat = IndAttType.End;
                lowPri = 1;
            }
            if (a.Name == W.firstLine)
            {
                iat = IndAttType.FirstLineOrHanging;
                lowPri = 4;
            }
            if (a.Name == W.firstLineChars)
            {
                iat = IndAttType.FirstLineOrHanging;
                lowPri = 3;
            }
            if (a.Name == W.hanging)
            {
                iat = IndAttType.FirstLineOrHanging;
                lowPri = 2;
            }
            if (a.Name == W.hangingChars)
            {
                iat = IndAttType.FirstLineOrHanging;
                lowPri = 1;
            }
            if (iat == IndAttType.None)
                throw new OpenXmlPowerToolsException("Internal error");
            var rv = new IndAttInfo
            {
                Attribute = a,
                AttributeType = iat,
                HighPri = highPri,
                LowPri = lowPri,
            };
            return rv;
        }

        // merge child tab elements
        // they are additive, with the exception that if there are two elements at the same location,
        // we need to take the higher, and not take the lower.
        private static XElement TabsMerge(XElement higherPriorityElement, XElement lowerPriorityElement)
        {
            if (higherPriorityElement != null && lowerPriorityElement == null)
                return higherPriorityElement;
            if (higherPriorityElement == null && lowerPriorityElement != null)
                return lowerPriorityElement;
            if (higherPriorityElement == null && lowerPriorityElement == null)
                return null;
            var hps = higherPriorityElement.Elements().Select(e =>
                new
                {
                    Pos = (int)e.Attribute(W.pos),
                    Pri = 1,
                    Element = e,
                }
            );
            var lps = lowerPriorityElement.Elements().Select(e =>
                new
                {
                    Pos = (int)e.Attribute(W.pos),
                    Pri = 2,
                    Element = e,
                }
            );
            var newTabElements = hps.Concat(lps)
                .GroupBy(s => s.Pos)
                .Select(g => g.OrderBy(s => s.Pri).First().Element)
                .Where(e => (string)e.Attribute(W.val) != "clear")
                .OrderBy(e => (int)e.Attribute(W.pos));
            var newTabs = new XElement(W.tabs, newTabElements);
            return newTabs;
        }

        private static XElement SpacingMerge(XElement hn, XElement ln)
        {
            if (hn == null && ln == null)
                return null;
            if (hn != null && ln == null)
                return hn;
            if (hn == null && ln != null)
                return ln;
            var mn1 = new XElement(W.spacing,
                hn.Attributes(),
                ln.Attributes().Where(a => hn.Attribute(a.Name) == null));
            return mn1;
        }

        private static IEnumerable<XElement> TableStyleStack(WordprocessingDocument wDoc, string tblStyleName)
        {
            XDocument sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            string currentStyle = tblStyleName;
            while (true)
            {
                XElement style = sXDoc
                    .Root
                    .Elements(W.style).Where(s => (string)s.Attribute(W.type) == "table" &&
                        (string)s.Attribute(W.styleId) == currentStyle)
                    .FirstOrDefault();
                if (style == null)
                    yield break;
                yield return style;
                currentStyle = (string)style.Elements(W.basedOn).Attributes(W.val).FirstOrDefault();
                if (currentStyle == null)
                    yield break;
            }
        }

        private static void AnnotateParagraphsWithListItems(WordprocessingDocument wDoc, XElement rootElement)
        {
            XName upPr = PtOpenXml.pt + "pPr";
            foreach (var para in rootElement.Descendants(W.p))
            {
                var listItemInfo = para.Annotation<ListItemRetriever.ListItemInfo>();
                if (listItemInfo != null)
                {
                    if (listItemInfo.IsListItem)
                    {
                        XElement newParaProps = null;

                        XElement lipPr = listItemInfo.Lvl.Element(W.pPr);
                        if (lipPr != null)
                        {
                            XElement currentParaProps = para.Element(PtOpenXml.pt + "pPr");
                            XElement mergedParaProps = MergeStyleElement(lipPr, currentParaProps);
                            newParaProps = new XElement(upPr, mergedParaProps.Elements());
                            if (currentParaProps != null)
                            {
                                currentParaProps.ReplaceWith(newParaProps);
                            }
                            else
                            {
                                para.Add(newParaProps);
                            }
                        }

                        XElement lirPr = listItemInfo.Lvl.Element(W.rPr);
                        if (lirPr != null)
                        {
                            XElement currentParaRunProps = newParaProps.Element(W.rPr);
                            XElement mergedParaRunProps = MergeStyleElement(lirPr, currentParaRunProps);
                            if (currentParaRunProps != null)
                            {
                                currentParaRunProps.ReplaceWith(mergedParaRunProps);
                            }
                            else
                            {
                                newParaProps.AddFirst(mergedParaRunProps);
                            }
                        }
                    }
                }
            }
        }

        private static XElement FontMerge(XElement higherPriorityFont, XElement lowerPriorityFont)
        {
            XElement rFonts;

            if (higherPriorityFont == null)
                return lowerPriorityFont;
            if (lowerPriorityFont == null)
                return higherPriorityFont;
            if (higherPriorityFont == null && lowerPriorityFont == null)
                return null;

            rFonts = new XElement(W.rFonts,
                (higherPriorityFont.Attribute(W.ascii) != null || higherPriorityFont.Attribute(W.asciiTheme) != null) ?
                    new [] {higherPriorityFont.Attribute(W.ascii), higherPriorityFont.Attribute(W.asciiTheme)} :
                    new [] {lowerPriorityFont.Attribute(W.ascii), lowerPriorityFont.Attribute(W.asciiTheme)},
                (higherPriorityFont.Attribute(W.hAnsi) != null || higherPriorityFont.Attribute(W.hAnsiTheme) != null) ?
                    new [] {higherPriorityFont.Attribute(W.hAnsi), higherPriorityFont.Attribute(W.hAnsiTheme)} :
                    new [] {lowerPriorityFont.Attribute(W.hAnsi), lowerPriorityFont.Attribute(W.hAnsiTheme)},
                (higherPriorityFont.Attribute(W.eastAsia) != null || higherPriorityFont.Attribute(W.eastAsiaTheme) != null) ?
                    new [] {higherPriorityFont.Attribute(W.eastAsia), higherPriorityFont.Attribute(W.eastAsiaTheme)} :
                    new [] {lowerPriorityFont.Attribute(W.eastAsia), lowerPriorityFont.Attribute(W.eastAsiaTheme)},
                (higherPriorityFont.Attribute(W.cs) != null || higherPriorityFont.Attribute(W.cstheme) != null) ?
                    new [] {higherPriorityFont.Attribute(W.cs), higherPriorityFont.Attribute(W.cstheme)} :
                    new [] {lowerPriorityFont.Attribute(W.cs), lowerPriorityFont.Attribute(W.cstheme)},
                (higherPriorityFont.Attribute(W.hint) != null ? higherPriorityFont.Attribute(W.hint) :
                    lowerPriorityFont.Attribute(W.hint))
            );

            return rFonts;
        }

        private static void AnnotateParagraphs(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement root, FormattingAssemblerSettings settings)
        {
            foreach (var para in root.Descendants(W.p))
	        {
                AnnotateParagraph(fai, wDoc, para, settings);
	        }
        }

        private static void AnnotateParagraph(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement para, FormattingAssemblerSettings settings)
        {
            XElement localParaProps = para.Element(W.pPr);
            if (localParaProps == null) {
                localParaProps = new XElement(W.pPr);
            }

            // get para table props, to be merged.
            XElement tablepPr = null;

            var blockLevelContentContainer = para
                .Ancestors()
                .FirstOrDefault(a => a.Name == W.body ||
                    a.Name == W.tbl ||
                    a.Name == W.txbxContent ||
                    a.Name == W.ftr ||
                    a.Name == W.hdr ||
                    a.Name == W.footnote ||
                    a.Name == W.endnote);
            if (blockLevelContentContainer.Name == W.tbl)
            {
                XElement tbl = blockLevelContentContainer;
                XElement style = tbl.Element(PtOpenXml.pt + "style");
                XElement cellCnf = para.Ancestors(W.tc).Take(1).Elements(W.tcPr).Elements(W.cnfStyle).FirstOrDefault();
                XElement rowCnf = para.Ancestors(W.tr).Take(1).Elements(W.trPr).Elements(W.cnfStyle).FirstOrDefault();

                if (style != null)
                {
                    // roll up tblPr, trPr, and tcPr from within a specific style.
                    // add each of these to the table, in PowerTools namespace.
                    tablepPr = style.Element(W.pPr);
                    if (tablepPr == null)
                        tablepPr = new XElement(W.pPr);

                    foreach (var ot in TableStyleOverrideTypes)
                    {
                        XName attName = TableStyleOverrideXNameMap[ot];
                        if (attName == null ||
                            (cellCnf != null && cellCnf.Attribute(attName).ToBoolean() == true) ||
                            (rowCnf != null && rowCnf.Attribute(attName).ToBoolean() == true))
                        {
                            XElement o = style
                                .Elements(W.tblStylePr)
                                .Where(tsp => (string)tsp.Attribute(W.type) == ot)
                                .FirstOrDefault();
                            if (o != null)
                            {
                                XElement otpPr = o.Element(W.pPr);
                                tablepPr = MergeStyleElement(otpPr, tablepPr);
                            }
                        }
                    }
                }
            }
            XElement rolledParaProps = ParaStyleRollup(fai, wDoc, para);
            XElement toggledParaProps = MergeStyleElement(rolledParaProps, tablepPr);
            XElement mergedParaProps = MergeStyleElement(localParaProps, toggledParaProps);
            string li = ListItemRetriever.RetrieveListItem(wDoc, para, null);
            ListItemRetriever.ListItemInfo lif = para.Annotation<ListItemRetriever.ListItemInfo>();
            if (lif != null && lif.IsListItem)
            {
                if (settings.RestrictToSupportedNumberingFormats)
                {
                    string numFmtForLevel = (string)lif.Lvl.Elements(W.numFmt).Attributes(W.val).FirstOrDefault();
                    if (numFmtForLevel == null)
                    {
                        var numFmtElement = lif.Lvl.Elements(MC.AlternateContent).Elements(MC.Choice).Elements(W.numFmt).FirstOrDefault();
                        if (numFmtElement != null && (string)numFmtElement.Attribute(W.val) == "custom")
                            numFmtForLevel = (string)numFmtElement.Attribute(W.format);
                    }
                    bool isLgl = lif.Lvl.Elements(W.isLgl).Any();
                    if (isLgl && numFmtForLevel != "decimalZero")
                        numFmtForLevel = "decimal";
                    if (!AcceptableNumFormats.Contains(numFmtForLevel))
                        throw new UnsupportedNumberingFormatException(numFmtForLevel + " is not a supported numbering format");
                }

                var lifInd = lif.Lvl.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
                var mppInd = mergedParaProps.Elements().FirstOrDefault(e => e.Name == W.ind);
                var paraInd = para.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
                var paraNumPr = para.Elements(W.pPr).Elements(W.numPr).FirstOrDefault();

                // The rule for priority for indentation:
                // if the w:numPr is in the paragraph, then the indentation from the numbering part takes precedence.
                // if the paragraph contains a style that contains a w:numPr, then the indentation from the style takes precedence.

                // Further, there is a rule:
                // if a paragraph contains a numPr with a numId=0, in other words, it is NOT a numbered item, then the indentation from the style
                // hierarchy is ignored.  This module does not follow this edge case rule.

                XElement ind = null;
                if (paraNumPr == null)
                    ind = IndMerge(mppInd, lifInd);
                else
                    ind = IndMerge(lifInd, mppInd);
                ind = IndMerge(paraInd, ind);
                mergedParaProps = new XElement(mergedParaProps.Name,
                    mergedParaProps.Attributes(),
                    mergedParaProps.Elements().Where(e => e.Name != W.ind),
                    ind);
            }

            XElement currentParaProps = para.Element(PtOpenXml.pt + "pPr");
            XElement newMergedParaProps = MergeStyleElement(mergedParaProps, currentParaProps);

            newMergedParaProps.Name = PtOpenXml.pt + "pPr";
            if (currentParaProps != null) {
                currentParaProps.ReplaceWith(newMergedParaProps);
            }
            else {
                para.Add(newMergedParaProps);
            }
        }

        private static string[] AcceptableNumFormats = new[] {
            "decimal",
            "decimalZero",
            "upperRoman",
            "lowerRoman",
            "upperLetter",
            "lowerLetter",
            "ordinal",
            "cardinalText",
            "ordinalText",
            "bullet",
            "0001, 0002, 0003, ...",
            "none",
        };

        private static XElement ParaStyleRollup(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement para)
        {
            var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var paraStyle = (string)para
                .Elements(W.pPr)
                .Elements(W.pStyle)
                .Attributes(W.val)
                .FirstOrDefault();
            if (paraStyle == null)
                paraStyle = fai.DefaultParagraphStyleName;
            var rolledUpParaStyleParaProps = new XElement(W.pPr);
            if (paraStyle != null) {
                rolledUpParaStyleParaProps = ParaStyleParaPropsStack(wDoc, paraStyle, para)
                    .Reverse()
                    .Aggregate(new XElement(W.pPr),
                        (r, s) => {
                            var newParaProps = MergeStyleElement(s, r);
                            return newParaProps;
                        });
            }
            return rolledUpParaStyleParaProps;
        }

        private static IEnumerable<XElement> ParaStyleParaPropsStack(WordprocessingDocument wDoc, string paraStyleName, XElement para)
        {
            var localParaStyleName = paraStyleName;
            var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            while (localParaStyleName != null) {
                XElement paraStyle = sXDoc.Root.Elements(W.style).FirstOrDefault(s =>
                    s.Attribute(W.type).Value == "paragraph" &&
                    s.Attribute(W.styleId).Value == localParaStyleName);
                if (paraStyle == null) {
                    yield break;
                }
                if (paraStyle.Element(W.pPr) == null) {
                    if (paraStyle.Element(W.rPr) != null)
                    {
                        var elementToYield2 = new XElement(W.pPr,
                            paraStyle.Element(W.rPr));
                        yield return elementToYield2;
                    }
                    localParaStyleName = (string)(paraStyle
                        .Elements(W.basedOn)
                        .Attributes(W.val)
                        .FirstOrDefault());
                    continue;
                }

                var elementToYield = new XElement(W.pPr,
                    paraStyle.Element(W.pPr).Attributes(),
                    paraStyle.Element(W.pPr).Elements(),
                    paraStyle.Element(W.rPr));
                yield return (elementToYield);

                var listItemInfo = para.Annotation<ListItemRetriever.ListItemInfo>();
                if (listItemInfo != null)
                {
                    if (listItemInfo.IsListItem)
                    {
                        XElement lipPr = listItemInfo.Lvl.Element(W.pPr);
                        if (lipPr == null)
                            lipPr = new XElement(W.pPr);
                        XElement lirPr = listItemInfo.Lvl.Element(W.rPr);
                        var elementToYield2 = new XElement(W.pPr,
                            lipPr.Attributes(),
                            lipPr.Elements(),
                            lirPr);
                        yield return (elementToYield2);
                    }
                }

                localParaStyleName = (string)paraStyle
                    .Elements(W.basedOn)
                    .Attributes(W.val)
                    .FirstOrDefault();
            }
            yield break;
        }

        private static void AnnotateRuns(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement root, FormattingAssemblerSettings settings)
        {
            var runsOrParas = root.Descendants()
                .Where(rp => {
                    return rp.Name == W.r || rp.Name == W.p;
                });
            foreach (var runOrPara in runsOrParas)
            {
                AnnotateRunProperties(fai, wDoc, runOrPara, settings);
            }
        }

        private static void AnnotateRunProperties(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement runOrPara, FormattingAssemblerSettings settings)
        {
            XElement localRunProps = null;
            if (runOrPara.Name == W.p) {
                var rPr = runOrPara.Elements(W.pPr).Elements(W.rPr).FirstOrDefault();
                if (rPr != null) {
                    localRunProps = rPr;
                }
            }
            else {
                localRunProps = runOrPara.Element(W.rPr);
            }
            if (localRunProps == null) {
                localRunProps = new XElement(W.rPr);
            }

            // get run table props, to be merged.
            XElement tablerPr = null;
            var blockLevelContentContainer = runOrPara
                .Ancestors()
                .FirstOrDefault(a => a.Name == W.body ||
                    a.Name == W.tbl ||
                    a.Name == W.txbxContent ||
                    a.Name == W.ftr ||
                    a.Name == W.hdr ||
                    a.Name == W.footnote ||
                    a.Name == W.endnote);
            if (blockLevelContentContainer.Name == W.tbl)
            {
                XElement tbl = blockLevelContentContainer;
                XElement style = tbl.Element(PtOpenXml.pt + "style");
                XElement cellCnf = runOrPara.Ancestors(W.tc).Take(1).Elements(W.tcPr).Elements(W.cnfStyle).FirstOrDefault();
                XElement rowCnf = runOrPara.Ancestors(W.tr).Take(1).Elements(W.trPr).Elements(W.cnfStyle).FirstOrDefault();

                if (style != null)
                {
                    tablerPr = style.Element(W.rPr);
                    if (tablerPr == null)
                        tablerPr = new XElement(W.rPr);

                    foreach (var ot in TableStyleOverrideTypes)
                    {
                        XName attName = TableStyleOverrideXNameMap[ot];
                        if (attName == null ||
                            (cellCnf != null && cellCnf.Attribute(attName).ToBoolean() == true) ||
                            (rowCnf != null && rowCnf.Attribute(attName).ToBoolean() == true))
                        {
                            XElement o = style
                                .Elements(W.tblStylePr)
                                .Where(tsp => (string)tsp.Attribute(W.type) == ot)
                                .FirstOrDefault();
                            if (o != null)
                            {
                                XElement otrPr = o.Element(W.rPr);
                                tablerPr = MergeStyleElement(otrPr, tablerPr);
                            }
                        }
                    }
                }
            }
            XElement rolledRunProps = CharStyleRollup(fai, wDoc, runOrPara);
            var toggledRunProps = ToggleMergeRunProps(rolledRunProps, tablerPr);
            var currentRunProps = runOrPara.Element(PtOpenXml.pt + "rPr"); // this is already stored on the run from previous aggregation of props
            var mergedRunProps = MergeStyleElement(toggledRunProps, currentRunProps);
            var newMergedRunProps = MergeStyleElement(localRunProps, mergedRunProps);
            AdjustFontAttributes(wDoc, runOrPara, newMergedRunProps, settings);

            newMergedRunProps.Name = PtOpenXml.pt + "rPr";
            if (currentRunProps != null) {
                currentRunProps.ReplaceWith(newMergedRunProps);
            }
            else {
                runOrPara.Add(newMergedRunProps);
            }
        }

        private static XElement CharStyleRollup(FormattingAssemblerInfo fai, WordprocessingDocument wDoc, XElement runOrPara)
        {
            var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();

            string charStyle = null;
            string paraStyle = null;
            XElement rPr = null;
            XElement pPr = null;
            XElement pStyle = null;
            XElement rStyle = null;
            CachedParaInfo cpi = null; // CachedParaInfo is an optimization for the case where a paragraph contains thousands of runs.

            if (runOrPara.Name == W.p)
            {
                cpi = runOrPara.Annotation<CachedParaInfo>();
                if (cpi != null)
                    pPr = cpi.ParagraphProperties;
                else
                {
                    pPr = runOrPara.Element(W.pPr);
                    if (pPr != null)
                    {
                        paraStyle = (string)pPr.Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                    }
                    else
                    {
                        paraStyle = fai.DefaultParagraphStyleName;
                    }
                    cpi = new CachedParaInfo
                    {
                        ParagraphProperties = pPr,
                        ParagraphStyleName = paraStyle,
                    };
                    runOrPara.AddAnnotation(cpi);
                }
                if (pPr != null) {
                    rPr = pPr.Element(W.rPr);
                }
            }
            else {
                rPr = runOrPara.Element(W.rPr);
            }
            if (rPr != null) {
                rStyle = rPr.Element(W.rStyle);
                if (rStyle != null)
                {
                    charStyle = (string)rStyle.Attribute(W.val);
                }
                else
                {
                    if (runOrPara.Name == W.r)
                        charStyle = (string)runOrPara
                            .Ancestors(W.p)
                            .Take(1)
                            .Elements(W.pPr)
                            .Elements(W.pStyle)
                            .Attributes(W.val)
                            .FirstOrDefault();
                }
            }

            if (charStyle == null)
            {
                if (runOrPara.Name == W.r)
                {
                    var ancestorPara = runOrPara.Ancestors(W.p).First();
                    cpi = ancestorPara.Annotation<CachedParaInfo>();
                    if (cpi != null)
                        charStyle = cpi.ParagraphStyleName;
                    else
                        charStyle = (string)runOrPara.Ancestors(W.p).First().Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                }
                if (charStyle == null)
                {
                    charStyle = fai.DefaultParagraphStyleName;
                }
            }

            // A run always must have an ancestor paragraph.
            XElement para = null;
            var rolledUpParaStyleRunProps = new XElement(W.rPr);
            if (runOrPara.Name == W.r) {
                para = runOrPara.Ancestors(W.p).FirstOrDefault();
            }
            else {
                para = runOrPara;
            }

            cpi = para.Annotation<CachedParaInfo>();
            if (cpi != null)
            {
                pPr = cpi.ParagraphProperties;
            }
            else
            {
                pPr = para.Element(W.pPr);
            }
            if (pPr != null) {
                pStyle = pPr.Element(W.pStyle);
                if (pStyle != null)
                {
                    paraStyle = (string)pStyle.Attribute(W.val);
                }
                else
                {
                    paraStyle = fai.DefaultParagraphStyleName;
                }
            }
            else
                paraStyle = fai.DefaultParagraphStyleName;

            string key = (paraStyle == null ? "[null]" : paraStyle) + "~|~" +
                (charStyle == null ? "[null]" : charStyle);
            XElement rolledRunProps = null;

            if (fai.RolledCharacterStyles.ContainsKey(key))
                rolledRunProps = fai.RolledCharacterStyles[key];
            else
            {
                XElement rolledUpCharStyleRunProps = new XElement(W.rPr);
                if (charStyle != null)
                {
                    rolledUpCharStyleRunProps =
                        CharStyleStack(wDoc, charStyle)
                            .Aggregate(new XElement(W.rPr),
                                (r, s) =>
                                {
                                    var newRunProps = MergeStyleElement(s, r);
                                    return newRunProps;
                                });
                }

                if (paraStyle != null)
                {
                    rolledUpParaStyleRunProps = ParaStyleRunPropsStack(wDoc, paraStyle)
                        .Aggregate(new XElement(W.rPr),
                            (r, s) =>
                            {
                                var newCharStyleRunProps = MergeStyleElement(s, r);
                                return newCharStyleRunProps;
                            });
                }
                rolledRunProps = MergeStyleElement(rolledUpCharStyleRunProps, rolledUpParaStyleRunProps);
                fai.RolledCharacterStyles.Add(key, rolledRunProps);
            }

            return rolledRunProps;
        }

        private static IEnumerable<XElement> ParaStyleRunPropsStack(WordprocessingDocument wDoc, string paraStyleName) 
        {
            var localParaStyleName = paraStyleName;
            var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var rValue = new Stack<XElement>();
            while (localParaStyleName != null) {
                var paraStyle = sXDoc.Root.Elements(W.style).FirstOrDefault(s => {
                    return (string)s.Attribute(W.type) == "paragraph" &&
                        (string)s.Attribute(W.styleId) == localParaStyleName;
                });
                if (paraStyle == null) {
                    return rValue;
                }
                if (paraStyle.Element(W.rPr) != null) {
                    rValue.Push(paraStyle.Element(W.rPr));
                }
                localParaStyleName = (string)paraStyle
                    .Elements(W.basedOn)
                    .Attributes(W.val)
                    .FirstOrDefault();
            }
            return rValue;
        }

        // returns collection of run properties
        private static IEnumerable<XElement> CharStyleStack(WordprocessingDocument wDoc, string charStyleName) 
        {
            var localCharStyleName = charStyleName;
            var sXDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var rValue = new Stack<XElement>();
            while (localCharStyleName != null) {
                XElement basedOn = null;
                // first look for character style
                var charStyle = sXDoc.Root.Elements(W.style).FirstOrDefault(s =>
                {
                    return (string)s.Attribute(W.type) == "character" &&
                        (string)s.Attribute(W.styleId) == localCharStyleName;
                });
                // if not found, look for paragraph style
                if (charStyle == null)
                {
                    charStyle = sXDoc.Root.Elements(W.style).FirstOrDefault(s =>
                    {
                        return (string)s.Attribute(W.styleId) == localCharStyleName;
                    });
                }
                if (charStyle == null) {
                    return rValue;
                }
                if (charStyle.Element(W.rPr) == null) {
                    basedOn = charStyle.Element(W.basedOn);
                    if (basedOn != null) {
                        localCharStyleName = (string)basedOn.Attribute(W.val);
                    }
                    else {
                        return rValue;
                    }
                }
                rValue.Push(charStyle.Element(W.rPr));
                localCharStyleName = null;
                basedOn = charStyle.Element(W.basedOn);
                if (basedOn != null) {
                    localCharStyleName = (string)basedOn.Attribute(W.val);
                }
            }
            return rValue;
        }

        private static XElement ToggleMergeRunProps(XElement higherPriorityElement, XElement lowerPriorityElement) 
        {
            if (lowerPriorityElement == null)
                return higherPriorityElement;
            if (higherPriorityElement == null)
                return lowerPriorityElement;

            var hpe = higherPriorityElement.Elements().Select(e => e.Name).ToArray();

            var newMergedElement = new XElement(higherPriorityElement.Name,
                higherPriorityElement.Attributes(),

                // process toggle properties
                higherPriorityElement.Elements()
                    .Where(e => { return e.Name != W.rFonts; })
                    .Select(higherChildElement => {
                        if (TogglePropertyNames.Contains(higherChildElement.Name)) {
                            var lowerChildElement = lowerPriorityElement.Element(higherChildElement.Name);
                            if (lowerChildElement == null) {
                                return higherChildElement;
                            }

                            var bHigher = higherChildElement.Attribute(W.val) == null || higherChildElement.Attribute(W.val).ToBoolean() == true;

                            var bLower = lowerChildElement.Attribute(W.val) == null || lowerChildElement.Attribute(W.val).ToBoolean() == true;

                            // if higher is true and lower is false, then return true element
                            if (bHigher && !bLower) {
                                return higherChildElement;
                            }

                            // if higher is false and lower is true, then return false element
                            if (!bHigher && bLower) {
                                return higherChildElement;
                            }

                            // if higher and lower are both true, then return false
                            if (bHigher && bLower) {
                                return new XElement(higherChildElement.Name,
                                    new XAttribute(W.val, "0"));
                            }

                            // otherwise, both higher and lower are false so can return higher element.
                            return higherChildElement;
                        }
                        return higherChildElement;
                    }),

                    FontMerge(higherPriorityElement.Element(W.rFonts), lowerPriorityElement.Element(W.rFonts)),

                    // take lower priority elements where there is not a higher priority element of same name
                    lowerPriorityElement.Elements()
                        .Where(e =>
                        {
                            return e.Name != W.rFonts && !hpe.Contains(e.Name);
                        }));

            return newMergedElement;
        }

        private static XName[] TogglePropertyNames = new [] {
            W.b,
            W.bCs,
            W.caps,
            W.emboss,
            W.i,
            W.iCs,
            W.imprint,
            W.outline,
            W.shadow,
            W.smallCaps,
            W.strike,
            W.vanish
        };

        private static XName[] PropertyNames = new [] {
            W.cs,
            W.rtl,
            W.u,
            W.color,
            W.highlight,
            W.shd
        };

        public class CharStyleAttributes
        {
            public string AsciiFont;
            public string HAnsiFont;
            public string EastAsiaFont;
            public string CsFont;
            public string Hint;
            public bool Rtl;

            public string LatinLang;
            public string BidiLang;
            public string EastAsiaLang;

            public Dictionary<XName, bool?> ToggleProperties;
            public Dictionary<XName, XElement> Properties;

            public CharStyleAttributes(XElement rPr)
            {
                ToggleProperties = new Dictionary<XName, bool?>();
                Properties = new Dictionary<XName, XElement>();

                if (rPr == null)
                    return;
                foreach (XName xn in TogglePropertyNames)
                {
                    ToggleProperties[xn] = GetBoolProperty(rPr, xn);
                }
                foreach (XName xn in PropertyNames)
                {
                    Properties[xn] = GetXmlProperty(rPr, xn);
                }
                var rFonts = rPr.Element(W.rFonts);
                if (rFonts == null)
                {
                    this.AsciiFont = null;
                    this.HAnsiFont = null;
                    this.EastAsiaFont = null;
                    this.CsFont = null;
                    this.Hint = null;
                }
                else
                {
                    this.AsciiFont = (string)(rFonts.Attribute(W.ascii));
                    this.HAnsiFont = (string)(rFonts.Attribute(W.hAnsi));
                    this.EastAsiaFont = (string)(rFonts.Attribute(W.eastAsia));
                    this.CsFont = (string)(rFonts.Attribute(W.cs));
                    this.Hint = (string)(rFonts.Attribute(W.hint));
                }
                XElement csel = this.Properties[W.cs];
                bool cs = csel != null && (csel.Attribute(W.val) == null || csel.Attribute(W.val).ToBoolean() == true);
                XElement rtlel = this.Properties[W.rtl];
                bool rtl = rtlel != null && (rtlel.Attribute(W.val) == null || rtlel.Attribute(W.val).ToBoolean() == true);
                Rtl = cs || rtl;
                var lang = rPr.Element(W.lang);
                if (lang != null)
                {
                    LatinLang = (string)lang.Attribute(W.val);
                    BidiLang = (string)lang.Attribute(W.bidi);
                    EastAsiaLang = (string)lang.Attribute(W.eastAsia);
                }
            }

            private static bool? GetBoolProperty(XElement rPr, XName propertyName)
            {
                if (rPr.Element(propertyName) == null)
                    return null;
                var s = (string)rPr.Element(propertyName).Attribute(W.val);
                if (s == null)
                    return true;
                if (s == "1") return true;
                if (s == "0") return false;
                if (s == "true") return true;
                if (s == "false") return false;
                if (s == "on") return true;
                if (s == "off") return false;
                return (bool)(rPr.Element(propertyName).Attribute(W.val));
            }

            private static XElement GetXmlProperty(XElement rPr, XName propertyName)
            {
                return rPr.Element(propertyName);
            }

            private static XName[] TogglePropertyNames = new[] {
                W.b,
                W.bCs,
                W.caps,
                W.emboss,
                W.i,
                W.iCs,
                W.imprint,
                W.outline,
                W.shadow,
                W.smallCaps,
                W.strike,
                W.vanish
            };

            private static XName[] PropertyNames = new[] {
                W.cs,
                W.rtl,
                W.u,
                W.color,
                W.highlight,
                W.shd
            };

        }

        private static void AdjustFontAttributes(WordprocessingDocument wDoc, XElement run, XElement rPr, FormattingAssemblerSettings settings) {
            XDocument themeXDoc = null;
            if (wDoc.MainDocumentPart.ThemePart != null)
                themeXDoc = wDoc.MainDocumentPart.ThemePart.GetXDocument();

            XElement fontScheme = null;
            XElement majorFont = null;
            XElement minorFont = null;
            if (themeXDoc != null)
            {
                fontScheme = themeXDoc.Root.Element(A.themeElements).Element(A.fontScheme);
                majorFont = fontScheme.Element(A.majorFont);
                minorFont = fontScheme.Element(A.minorFont);
            }
            var rFonts = rPr.Element(W.rFonts);
            if (rFonts == null) {
                return;
            }
            var asciiTheme = (string)rFonts.Attribute(W.asciiTheme);
            var hAnsiTheme = (string)rFonts.Attribute(W.hAnsiTheme);
            var eastAsiaTheme = (string)rFonts.Attribute(W.eastAsiaTheme);
            var cstheme = (string)rFonts.Attribute(W.cstheme);
            string ascii = null;
            string hAnsi = null;
            string eastAsia = null;
            string cs = null;

            XElement minorLatin = null;
            string minorLatinTypeface = null;
            XElement majorLatin = null;
            string majorLatinTypeface = null;

            if (minorFont != null)
            {
                minorLatin = minorFont.Element(A.latin);
                minorLatinTypeface = (string)minorLatin.Attribute("typeface");
            }

            if (majorFont != null)
            {
                majorLatin = majorFont.Element(A.latin);
                majorLatinTypeface = (string)majorLatin.Attribute("typeface");
            }
            if (asciiTheme != null) {
                if (asciiTheme.StartsWith("minor") && minorLatinTypeface != null) {
                    ascii = minorLatinTypeface;
                }
                else if (asciiTheme.StartsWith("major") && majorLatinTypeface != null) {
                    ascii = majorLatinTypeface;
                }
            }
            if (hAnsiTheme != null) {
                if (hAnsiTheme.StartsWith("minor") && minorLatinTypeface != null) {
                    hAnsi = minorLatinTypeface;
                }
                else if (hAnsiTheme.StartsWith("major") && majorLatinTypeface != null) {
                    hAnsi = majorLatinTypeface;
                }
            }
            if (eastAsiaTheme != null) {
                if (eastAsiaTheme.StartsWith("minor") && minorLatinTypeface != null) {
                    eastAsia = minorLatinTypeface;
                }
                else if (eastAsiaTheme.StartsWith("major") && majorLatinTypeface != null) {
                    eastAsia = majorLatinTypeface;
                }
            }
            if (cstheme != null) {
                if (cstheme.StartsWith("minor") && minorFont != null) {
                    cs = (string)minorFont.Element(A.cs).Attribute("typeface");
                }
                else if (cstheme.StartsWith("major") && majorFont != null) {
                    cs = (string)majorFont.Element(A.cs).Attribute("typeface");
                }
            }

            if (ascii != null) {
                rFonts.SetAttributeValue(W.ascii, ascii);
            }
            if (hAnsi != null) {
                rFonts.SetAttributeValue(W.hAnsi, hAnsi);
            }
            if (eastAsia != null) {
                rFonts.SetAttributeValue(W.eastAsia, eastAsia);
            }
            if (cs != null) {
                rFonts.SetAttributeValue(W.cs, cs);
            }

            var str = run.Elements(W.t).Select(t => { return (string)t; }).StringConcatenate();
            if (str.Length == 0)
                return;   // no 'FontFamily' annotation
            var csa = new CharStyleAttributes(rPr);

            // This module determines the font based on just the first character.
            // Technically, a run can contain characters from different Unicode code blocks, and hence should be rendered with different fonts.
            // However, Word breaks up runs that use more than one font into multiple runs.  Other producers of WordprocessingML may not, so in
            // that case, this routine may need to be augmented to look at all characters in a run.

            /*
            old code
            var fontFamilies = str.select(function (c) {
                var ft = Pav.DetermineFontTypeFromCharacter(c, csa);
                switch (ft) {
                    case Pav.FontType.Ascii:
                        return cast(rFonts.attribute(W.ascii));
                    case Pav.FontType.HAnsi:
                        return cast(rFonts.attribute(W.hAnsi));
                    case Pav.FontType.EastAsia:
                        return cast(rFonts.attribute(W.eastAsia));
                    case Pav.FontType.CS:
                        return cast(rFonts.attribute(W.cs));
                    default:
                        return null;
                }
            })
                .where(function (f) { return f != null && f != ""; })
                .distinct()
                .select(function (f) { return new Pav.FontFamily(f); })
                .toArray();
            */

            var ft = DetermineFontTypeFromCharacter(str[0], csa);
            string fontType = null;
            switch (ft) {
                case FontType.Ascii:
                    fontType = (string)rFonts.Attribute(W.ascii);
                    break;
                case FontType.HAnsi:
                    fontType = (string)rFonts.Attribute(W.hAnsi);
                    break;
                case FontType.EastAsia:
                    if (settings.RestrictToSupportedLanguages)
                        throw new UnsupportedLanguageException("EastAsia languages are not supported");
                    fontType = (string)rFonts.Attribute(W.eastAsia);
                    break;
                case FontType.CS:
                    if (settings.RestrictToSupportedLanguages)
                        throw new UnsupportedLanguageException("Complex script (RTL) languages are not supported");
                    fontType = (string)rFonts.Attribute(W.cs);
                    break;
            }

            XAttribute fta = new XAttribute(PtOpenXml.pt + "FontName", fontType.ToString());
            run.Add(fta);
        }

        public enum FontType
        {
            Ascii,
            HAnsi,
            EastAsia,
            CS
        };
        
        // The algorithm for this method comes from the implementer notes in [MS-OI29500].pdf
        // section 2.1.87

        // The implementer notes are at:
        // http://msdn.microsoft.com/en-us/library/ee908652.aspx

        public static FontType DetermineFontTypeFromCharacter(char ch, CharStyleAttributes csa)
        {
            // If the run has the cs element ("[ISO/IEC-29500-1] §17.3.2.7; cs") or the rtl element ("[ISO/IEC-29500-1] §17.3.2.30; rtl"),
            // then the cs (or cstheme if defined) font is used, regardless of the Unicode character values of the run’s content.
            if (csa.Rtl)
            {
                return FontType.CS;
            }

            // A large percentage of characters will fall in the following rule.

            // Unicode Block: Basic Latin
            if (ch >= 0x00 && ch <= 0x7f)
            {
                return FontType.Ascii;
            }

            // If the eastAsia (or eastAsiaTheme if defined) attribute’s value is “Times New Roman” and the ascii (or asciiTheme if defined)
            // and hAnsi (or hAnsiTheme if defined) attributes are equal, then the ascii (or asciiTheme if defined) font is used.
            if (csa.EastAsiaFont == "Times New Roman" &&
                csa.AsciiFont == csa.HAnsiFont)
            {
                return FontType.Ascii;
            }

            // Unicode BLock: Latin-1 Supplement
            if (ch >= 0xA0 && ch <= 0xFF)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (ch == 0xA1 ||
                        ch == 0xA4 ||
                        ch == 0xA7 ||
                        ch == 0xA8 ||
                        ch == 0xAA ||
                        ch == 0xAD ||
                        ch == 0xAF ||
                        (ch >= 0xB0 && ch <= 0xB4) ||
                        (ch >= 0xB6 && ch <= 0xBA) ||
                        (ch >= 0xBC && ch <= 0xBF) ||
                        ch == 0xD7 ||
                        ch == 0xF7)
                    {
                        return FontType.EastAsia;
                    }
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans")
                    {
                        if (ch == 0xE0 ||
                            ch == 0xE1 ||
                            (ch >= 0xE8 && ch <= 0xEA) ||
                            (ch >= 0xEC && ch <= 0xED) ||
                            (ch >= 0xF2 && ch <= 0xF3) ||
                            (ch >= 0xF9 && ch <= 0xFA) ||
                            ch == 0xFC)
                        {
                            return FontType.EastAsia;
                        }
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Latin Extended-A
            if (ch >= 0x0100 && ch <= 0x017F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Latin Extended-B
            if (ch >= 0x0180 && ch <= 0x024F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: IPA Extensions
            if (ch >= 0x0250 && ch <= 0x02AF)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (csa.EastAsiaLang == "zh-hant" ||
                        csa.EastAsiaLang == "zh-hans"
                        /* || the character set of the east Asia (or east Asia theme) font is Chinese5 || GB2312 todo */)
                    {
                        return FontType.EastAsia;
                    }
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Spacing Modifier Letters
            if (ch >= 0x02B0 && ch <= 0x02FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Combining Diacritic Marks
            if (ch >= 0x0300 && ch <= 0x036F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Greek
            if (ch >= 0x0370 && ch <= 0x03CF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Cyrillic
            if (ch >= 0x0400 && ch <= 0x04FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Hebrew
            if (ch >= 0x0590 && ch <= 0x05FF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Arabic
            if (ch >= 0x0600 && ch <= 0x06FF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Syriac
            if (ch >= 0x0700 && ch <= 0x074F)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Arabic Supplement
            if (ch >= 0x0750 && ch <= 0x077F)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Thanna
            if (ch >= 0x0780 && ch <= 0x07BF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Hangul Jamo
            if (ch >= 0x1100 && ch <= 0x11FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Latin Extended Additional
            if (ch >= 0x1E00 && ch <= 0x1EFF)
            {
                if (csa.Hint == "eastAsia" &&
                    (csa.EastAsiaLang == "zh-hant" ||
                    csa.EastAsiaLang == "zh-hans"))
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: General Punctuation
            if (ch >= 0x2000 && ch <= 0x206F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Superscripts and Subscripts
            if (ch >= 0x2070 && ch <= 0x209F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Currency Symbols
            if (ch >= 0x20A0 && ch <= 0x20CF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Combining Diacritical Marks for Symbols
            if (ch >= 0x20D0 && ch <= 0x20FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Letter-like Symbols
            if (ch >= 0x2100 && ch <= 0x214F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Number Forms
            if (ch >= 0x2150 && ch <= 0x218F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Arrows
            if (ch >= 0x2190 && ch <= 0x21FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Mathematical Operators
            if (ch >= 0x2200 && ch <= 0x22FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Miscellaneous Technical
            if (ch >= 0x2300 && ch <= 0x23FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Control Pictures
            if (ch >= 0x2400 && ch <= 0x243F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Optical Character Recognition
            if (ch >= 0x2440 && ch <= 0x245F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Enclosed Alphanumerics
            if (ch >= 0x2460 && ch <= 0x24FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Box Drawing
            if (ch >= 0x2500 && ch <= 0x257F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Block Elements
            if (ch >= 0x2580 && ch <= 0x259F)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Geometric Shapes
            if (ch >= 0x25A0 && ch <= 0x25FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Miscellaneous Symbols
            if (ch >= 0x2600 && ch <= 0x26FF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Dingbats
            if (ch >= 0x2700 && ch <= 0x27BF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: CJK Radicals Supplement
            if (ch >= 0x2E80 && ch <= 0x2EFF)
            {
                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Kangxi Radicals
            if (ch >= 0x2F00 && ch <= 0x2FDF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Ideographic Description Characters
            if (ch >= 0x2FF0 && ch <= 0x2FFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Symbols and Punctuation
            if (ch >= 0x3000 && ch <= 0x303F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hiragana
            if (ch >= 0x3040 && ch <= 0x309F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Katakana
            if (ch >= 0x30A0 && ch <= 0x30FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Bopomofo
            if (ch >= 0x3100 && ch <= 0x312F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hangul Compatibility Jamo
            if (ch >= 0x3130 && ch <= 0x318F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Kanbun
            if (ch >= 0x3190 && ch <= 0x319F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Enclosed CJK Letters and Months
            if (ch >= 0x3200 && ch <= 0x32FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Compatibility
            if (ch >= 0x3300 && ch <= 0x33FF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Unified Ideographs Extension A
            if (ch >= 0x3400 && ch <= 0x4DBF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: CJK Unified Ideographs
            if (ch >= 0x4E00 && ch <= 0x9FAF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Yi Syllables
            if (ch >= 0xA000 && ch <= 0xA48F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Yi Radicals
            if (ch >= 0xA490 && ch <= 0xA4CF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Hangul Syllables
            if (ch >= 0xAC00 && ch <= 0xD7AF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: High Surrogates
            if (ch >= 0xD800 && ch <= 0xDB7F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: High Private Use Surrogates
            if (ch >= 0xDB80 && ch <= 0xDBFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Low Surrogates
            if (ch >= 0xDC00 && ch <= 0xDFFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Private Use Area
            if (ch >= 0xE000 && ch <= 0xF8FF)
            {
                // per the standard, it says that E713 should be EastAsia only if hint == "eastAsia"
                // however, per Word it is always eastAsia.
                //return FontType.EastAsia;

                if (csa.Hint == "eastAsia")
                {
                    return FontType.EastAsia;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: CJK Compatibility Ideographs
            if (ch >= 0xF900 && ch <= 0xFAFF)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Alphabetic Presentation Forms
            if (ch >= 0xFB00 && ch <= 0xFB4F)
            {
                if (csa.Hint == "eastAsia")
                {
                    if (ch >= 0xFB00 && ch <= 0xFB1C)
                        return FontType.EastAsia;
                    if (ch >= 0xFB1D && ch <= 0xFB4F)
                        return FontType.Ascii;
                }
                return FontType.HAnsi;
            }

            // Unicode Block: Arabic Presentation Forms-A
            if (ch >= 0xFB50 && ch <= 0xFDFF)
            {
                return FontType.Ascii;
            }

            // Unicode Block: CJK Compatibility Forms
            if (ch >= 0xFE30 && ch <= 0xFE4F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Small Form Variants
            if (ch >= 0xFE50 && ch <= 0xFE6F)
            {
                return FontType.EastAsia;
            }

            // Unicode Block: Arabic Presentation Forms-B
            if (ch >= 0xFE70 && ch <= 0xFEFE)
            {
                return FontType.Ascii;
            }

            // Unicode Block: Halfwidth and Fullwidth Forms
            if (ch >= 0xFF00 && ch <= 0xFFEF)
            {
                return FontType.EastAsia;
            }
            return FontType.HAnsi;
        }

        private static bool? GetBoolProperty(XElement rPr, XName propertyName)
        {
            if (rPr.Element(propertyName) == null) {
                return null;
            }
            var property = rPr.Element(propertyName).Attribute(W.val);
            if (property == null)
                return true;
            return property.ToBoolean();
        }

        private static XElement GetXmlProperty(XElement rPr, XName propertyName)
        {
            return rPr.Element(propertyName);
        }

        private class FormattingAssemblerInfo
        {
            public string DefaultParagraphStyleName;
            public string DefaultCharacterStyleName;
            public string DefaultTableStyleName;
            public Dictionary<string, XElement> RolledCharacterStyles;
            public FormattingAssemblerInfo()
            {
                RolledCharacterStyles = new Dictionary<string, XElement>();
            }
        }

        // CachedParaInfo is an optimization for the case where a paragraph contains thousands of runs.
        private class CachedParaInfo
        {
            public string ParagraphStyleName;
            public XElement ParagraphProperties;
        }

        public class UnsupportedNumberingFormatException : Exception
        {
            public UnsupportedNumberingFormatException(string message) : base(message) { }
        }

        public class UnsupportedLanguageException : Exception
        {
            public UnsupportedLanguageException(string message) : base(message) { }
        }
    }
}
