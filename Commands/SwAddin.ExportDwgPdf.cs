using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    /// <summary>
    /// 图纸转档功能：导出工程图为 DWG 和 PDF
    /// </summary>
    public partial class SwAddin
    {
        /// <summary>
        /// 启用条件：仅在打开工程图时可用（否则按钮变灰）
        /// </summary>
        public int EnableExportDwgPdf()
        {
            if (swApp.ActiveDoc == null) return 0;
            ModelDoc2 modDoc = (ModelDoc2)swApp.ActiveDoc;
            return (modDoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING) ? 1 : 0;
        }

        // 回调方法：导出 DWG/PDF
        [ComVisible(true)]
        public void ExportDwgPdf()
        {

            try
            {
                if (!(swApp.ActiveDoc is IModelDoc2 doc))
                {
                    MessageBox.Show("请先打开一个文档。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (doc.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("仅支持工程图（.slddrw）导出DWG/PDF。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string fileName = doc.GetPathName();
                if (string.IsNullOrEmpty(fileName))
                {
                    MessageBox.Show("当前文档未保存，请先保存。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 导出前：图面布满显示（整屏 / 适应窗口）
                ZoomDrawingToFit(doc);

                string shortName = Path.ChangeExtension(fileName, null);
                string dwgName = shortName + ".DWG";
                string pdfName = shortName + ".PDF";

                DeleteFileIfExists(dwgName);
                DeleteFileIfExists(pdfName);

                int errors = 0, warnings = 0;

                doc.Extension.SaveAs(
                    dwgName,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

                doc.Extension.SaveAs(
                    pdfName,
                    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                    null,
                    ref errors,
                    ref warnings);

            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message, "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteFileIfExists(string filePath)
        {
            if (File.Exists(filePath))
            {
                try
                {
                    File.SetAttributes(filePath, FileAttributes.Normal);
                    File.Delete(filePath);
                }
                catch (Exception)
                {}
            }
        }
    }
}
