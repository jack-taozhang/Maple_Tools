using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    /// <summary>
    /// 模型更新功能：从 Excel 读取参数，按组件层级从深到浅更新方程式
    /// 使用无依赖 ExcelReader 读取 .xlsx（无需安装 Office）
    /// </summary>
    public partial class SwAddin
    {
        // 回调方法：从 Excel 读取参数更新模型方程式
        [ComVisible(true)]
        public void UpdateModelsFromExcel()
        {

            try
            {
                if (!(swApp.ActiveDoc is IModelDoc2 swModel))
                {
                    MessageBox.Show("请先打开SolidWorks文档", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string modelPath = swModel.GetPathName();
                if (string.IsNullOrEmpty(modelPath))
                {
                    MessageBox.Show("请先保存文档", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 记录最初打开的文档
                originallyOpenedDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                DocumentOps.RecordOpenedDocs(swApp, originallyOpenedDocs);
                originalOpenDocsPath = modelPath;

                // 查找 Excel 文件（与模型同目录的"Design dimensions.xlsx"）
                string folderPath = Path.GetDirectoryName(modelPath);
                string excelPath = Path.Combine(folderPath, "Design dimensions.xlsx");
                if (!File.Exists(excelPath))
                {
                    MessageBox.Show("Excel文件不存在: " + excelPath, "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 读取 Excel 参数（"设计数据"工作表）
                var paramDict = ExcelReader.ReadParams(excelPath, "设计数据");
                if (paramDict.Count == 0)
                {
                    MessageBox.Show("Excel文件中未找到有效参数", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 初始化层级结构
                visitedAssemblies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                componentHierarchy = new List<ComponentInfo>();
                processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                int docType = swModel.GetType();
                if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    if (swModel is IAssemblyDoc swAssembly)
                    {
                        BuildComponentHierarchy(swApp, swAssembly, 0);
                        ProcessByHierarchyForEquations(paramDict);
                    }
                }
                else if (docType == (int)swDocumentTypes_e.swDocPART)
                {
                    UpdateDocumentEquations(swModel, paramDict);
                }

                // 关闭所有新打开的文档（非最初打开的）
                CloseNonOriginallyOpenedDocs();

                // 重新激活原始文档
                if (!(swApp.ActiveDoc is IModelDoc2 activeDoc &&
                      string.Equals(activeDoc.GetPathName(), originalOpenDocsPath, StringComparison.OrdinalIgnoreCase)))
                {
                    int err = 0;
                    try { swApp.ActivateDoc3(originalOpenDocsPath, false, 1, ref err); }
                    catch (Exception) {}
                }

                MessageBox.Show("模型根据文件参数已更新完成!", "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) {
                if (originallyOpenedDocs != null)
                {
                    try { CloseNonOriginallyOpenedDocs(); }
                    catch { }
                }
                MessageBox.Show("发生错误: " + ex.Message, "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 按层级从深到浅处理组件（更新方程式版本）
        private void ProcessByHierarchyForEquations(Dictionary<string, double> paramDict)
        {
            if (componentHierarchy.Count == 0) return;

            int maxLevel = 0;
            foreach (var comp in componentHierarchy)
            {
                if (comp.Level > maxLevel) maxLevel = comp.Level;
            }

            for (int i = maxLevel; i >= 0; i--)
            {
                foreach (var comp in componentHierarchy)
                {
                    if (comp.Level != i) continue;
                    if (processedFiles.Contains(comp.Path)) continue;
                    processedFiles.Add(comp.Path);

                    bool wasOriginallyOpened = originallyOpenedDocs.Contains(comp.Path);

                    IModelDoc2 swModel = DocumentOps.GetOpenedDocument(swApp, comp.Path)
                        ?? DocumentOps.OpenDocumentSilent(swApp, comp.Path);

                    if (swModel != null)
                    {
                        UpdateDocumentEquations(swModel, paramDict);

                        if (!wasOriginallyOpened)
                        {
                            try { swApp.CloseDoc(comp.Path); }
                            catch (Exception) {}
                        }
                    }
                }
            }
        }

        // 更新文档中的方程式
        private void UpdateDocumentEquations(IModelDoc2 swModel, Dictionary<string, double> paramDict)
        {
            string modelPath = swModel.GetPathName();

            if (!(swApp.ActiveDoc is IModelDoc2 activeDoc &&
                  string.Equals(activeDoc.GetPathName(), modelPath, StringComparison.OrdinalIgnoreCase)))
            {
                int err = 0;
                try { swApp.ActivateDoc3(modelPath, false, 1, ref err); }
                catch (Exception) {}
            }

            IEquationMgr swEqMgr;
            try
            {
                swEqMgr = swModel.GetEquationMgr() as IEquationMgr;
            }
            catch (Exception) {
                return;
            }
            if (swEqMgr == null) return;

            try { swModel.FeatureManager.EnableFeatureTree = false; }
            catch { }

            foreach (var kvp in paramDict)
            {
                ModifyEquation(swEqMgr, kvp.Key, kvp.Value);
            }

            try { swModel.FeatureManager.EnableFeatureTree = true; }
            catch { }
            try { swModel.EditRebuild3(); }
            catch (Exception) {}
            try { swModel.Save(); }
            catch (Exception) {}

            if (!string.Equals(originalOpenDocsPath, modelPath, StringComparison.OrdinalIgnoreCase))
            {
                try { swApp.CloseDoc(modelPath); }
                catch (Exception) {}
            }
        }

        // 修改方程式中指定变量的值（保留复杂表达式，只替换简单数值）
        private bool ModifyEquation(IEquationMgr eqMgr, string targetVar, double newVal)
        {
            try
            {
                int count = eqMgr.GetCount();
                for (int i = 0; i < count; i++)
                {
                    string eqText = eqMgr.get_Equation(i);
                    if (string.IsNullOrEmpty(eqText)) continue;

                    int eqIdx = eqText.IndexOf('=');
                    if (eqIdx < 0) continue;

                    string leftPart = eqText.Substring(0, eqIdx).Trim();
                    string expr = eqText.Substring(eqIdx + 1).Trim();

                    string varName = leftPart.Replace("\"", "").Trim();

                    if (string.Equals(varName, targetVar, StringComparison.OrdinalIgnoreCase))
                    {
                        bool isComplex = expr.Contains("(") || expr.Contains("+") ||
                            expr.Contains("-") || expr.Contains("*") || expr.Contains("/");

                        if (!isComplex)
                        {
                            eqMgr.set_Equation(i, "\"" + targetVar + "\" = " + newVal.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {}
            return false;
        }

        // 关闭所有非最初打开的文档
        private void CloseNonOriginallyOpenedDocs()
        {
            try
            {
                HashSet<string> currentDocs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                DocumentOps.RecordOpenedDocs(swApp, currentDocs);

                foreach (string docPath in currentDocs)
                {
                    if (!originallyOpenedDocs.Contains(docPath))
                    {
                        try { swApp.CloseDoc(docPath); }
                        catch (Exception) {}
                    }
                }
            }
            catch (Exception)
            {}
        }
    }
}
