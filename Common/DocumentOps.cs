using System;
using System.IO;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    /// <summary>
    /// SolidWorks 文档操作工具（静态）
    /// </summary>
    internal static class DocumentOps
    {
        // 静默打开文档
        internal static IModelDoc2 OpenDocumentSilent(ISldWorks swApp, string filePath)
        {
            int docType;
            string ext = Path.GetExtension(filePath).ToLower();

            if (ext == ".sldasm")
                docType = (int)swDocumentTypes_e.swDocASSEMBLY;
            else
                docType = (int)swDocumentTypes_e.swDocPART;

            int errors = 0, warnings = 0;
            try
            {
                object doc = swApp.OpenDoc6(filePath, docType, 1, "", ref errors, ref warnings);
                return doc as IModelDoc2;
            }
            catch (Exception) {
                return null;
            }
        }

        // 获取已打开的文档
        internal static IModelDoc2 GetOpenedDocument(ISldWorks swApp, string filePath)
        {
            object docsObj = swApp.GetDocuments();
            if (docsObj == null) return null;

            if (!(docsObj is object[] docs)) return null;

            foreach (object doc in docs)
            {
                if (doc is IModelDoc2 model &&
                    string.Equals(model.GetPathName(), filePath, StringComparison.OrdinalIgnoreCase))
                    return model;
            }
            return null;
        }

        // 检查是否为 SpeedPak 组件
        internal static bool IsSpeedPak(IComponent2 swComp)
        {
            try
            {
                return swComp.IsSpeedPak;
            }
            catch
            {
                return false;
            }
        }

        // 记录当前所有打开的文档到集合
        internal static void RecordOpenedDocs(ISldWorks swApp, System.Collections.Generic.HashSet<string> dict)
        {
            try
            {
                object docsObj = swApp.GetDocuments();
                if (docsObj == null) return;
                if (!(docsObj is object[] docs)) return;
                foreach (object doc in docs)
                {
                    if (doc is IModelDoc2 model)
                    {
                        string p = model.GetPathName();
                        if (!string.IsNullOrEmpty(p)) dict.Add(p);
                    }
                }
            }
            catch (Exception)
            {}
        }
    }
}
