using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Mtools
{
    /// <summary>
    /// 无依赖 Excel 读取（用 System.IO.Compression 解析 .xlsx Open XML）
    /// 不依赖 Microsoft Office / Excel 应用程序
    /// </summary>
    internal static class ExcelReader
    {
        // 读取 Excel 参数：A列=变量名，B列=数值，从第2行开始（跳过表头）
        // sheetName 指定工作表名称（如"设计数据"）
        internal static Dictionary<string, double> ReadParams(string excelPath, string sheetName)
        {
            var paramDict = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    // 1. 读取共享字符串表
                    List<string> sharedStrings = ReadSharedStrings(zip);

                    // 2. 找到指定工作表对应的 XML 路径（找不到时回退 sheet1）
                    string sheetPath = FindWorksheetPath(zip, sheetName) ?? "xl/worksheets/sheet1.xml";

                    ZipArchiveEntry sheetEntry = zip.GetEntry(sheetPath);
                    if (sheetEntry == null)
                    {
                        return paramDict;
                    }

                    // 3. 解析工作表数据
                    using (var stream = sheetEntry.Open())
                    {
                        var doc = new XmlDocument();
                        doc.Load(stream);
                        var ns = new XmlNamespaceManager(doc.NameTable);
                        ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                        XmlNodeList rows = doc.SelectNodes("//s:sheetData/s:row", ns);
                        if (rows == null) return paramDict;

                        foreach (XmlNode rowNode in rows)
                        {
                            // 跳过第一行（表头）
                            XmlAttribute rowAttr = rowNode.Attributes["r"];
                            if (rowAttr != null)
                            {
                                if (int.TryParse(rowAttr.Value, out int rowNum) && rowNum < 2) continue;
                            }

                            XmlNodeList cells = rowNode.SelectNodes("s:c", ns);
                            if (cells == null || cells.Count < 2) continue;

                            string colA = GetCellText(cells[0], sharedStrings, ns);
                            double colB = GetCellNumber(cells[1], ns);

                            if (!string.IsNullOrWhiteSpace(colA))
                            {
                                paramDict[colA.Trim()] = colB;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {}
            return paramDict;
        }

        // 读取共享字符串表 xl/sharedStrings.xml
        private static List<string> ReadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            try
            {
                ZipArchiveEntry entry = zip.GetEntry("xl/sharedStrings.xml");
                if (entry == null) return list;

                using (var stream = entry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);
                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                    XmlNodeList siNodes = doc.SelectNodes("//s:si", ns);
                    if (siNodes == null) return list;

                    foreach (XmlNode siNode in siNodes)
                    {
                        XmlNode tNode = siNode.SelectSingleNode(".//s:t", ns);
                        list.Add(tNode != null ? tNode.InnerText : "");
                    }
                }
            }
            catch (Exception)
            {}
            return list;
        }

        // 根据工作表名称查找对应的 worksheet XML 路径
        private static string FindWorksheetPath(ZipArchive zip, string sheetName)
        {
            try
            {
                ZipArchiveEntry wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null) return null;

                string targetRId = null;
                using (var stream = wbEntry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);
                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                    ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                    XmlNodeList sheetNodes = doc.SelectNodes("//s:sheet", ns);
                    if (sheetNodes == null) return null;

                    foreach (XmlNode sheetNode in sheetNodes)
                    {
                        XmlAttribute nameAttr = sheetNode.Attributes["name"];
                        if (nameAttr != null && string.Equals(nameAttr.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            XmlAttribute idAttr = sheetNode.Attributes["id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"];
                            if (idAttr != null)
                            {
                                targetRId = idAttr.Value;
                                break;
                            }
                        }
                    }
                }

                if (targetRId == null) return null;

                ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relsEntry == null) return null;

                using (var stream = relsEntry.Open())
                {
                    var doc = new XmlDocument();
                    doc.Load(stream);
                    var ns = new XmlNamespaceManager(doc.NameTable);
                    ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

                    XmlNodeList relNodes = doc.SelectNodes("//r:Relationship", ns);
                    if (relNodes == null) return null;

                    foreach (XmlNode relNode in relNodes)
                    {
                        XmlAttribute idAttr = relNode.Attributes["Id"];
                        XmlAttribute targetAttr = relNode.Attributes["Target"];
                        if (idAttr != null && targetAttr != null &&
                            string.Equals(idAttr.Value, targetRId, StringComparison.OrdinalIgnoreCase))
                        {
                            string target = targetAttr.Value.Replace('\\', '/');
                            if (target.StartsWith("/"))
                                target = target.TrimStart('/');
                            else
                                target = "xl/" + target;
                            return target;
                        }
                    }
                }
            }
            catch (Exception)
            {}
            return null;
        }

        // 获取单元格文本值（处理字符串类型 t="s" 引用共享字符串）
        private static string GetCellText(XmlNode cellNode, List<string> sharedStrings, XmlNamespaceManager ns)
        {
            XmlAttribute typeAttr = cellNode.Attributes["t"];
            string type = (typeAttr != null) ? typeAttr.Value : "n";

            XmlNode vNode = cellNode.SelectSingleNode("s:v", ns);
            string v = (vNode != null) ? vNode.InnerText : "";

            if (type == "s" && sharedStrings.Count > 0)
            {
                if (int.TryParse(v, out int idx) && idx >= 0 && idx < sharedStrings.Count)
                    return sharedStrings[idx];
                return "";
            }
            return v;
        }

        // 获取单元格数值
        private static double GetCellNumber(XmlNode cellNode, XmlNamespaceManager ns)
        {
            XmlNode vNode = cellNode.SelectSingleNode("s:v", ns);
            if (vNode == null) return 0;

            if (double.TryParse(vNode.InnerText, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return 0;
        }
    }
}
