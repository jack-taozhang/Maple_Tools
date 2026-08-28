using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    /// <summary>
    /// 显示控制功能：批量设置零件/装配体的显示选项（基准面、轴、草图等）
    /// </summary>
    public partial class SwAddin
    {
        // 回调方法：应用显示设置到所有文档
        [ComVisible(true)]
        public void ControlSolidWorksDisplay()
        {

            try
            {
                if (!(swApp.ActiveDoc is IModelDoc2 swModel))
                {
                    MessageBox.Show("请先打开SolidWorks文档", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var displaySettings = CreateDisplaySettings();

                originalOpenDocsPath = swModel.GetPathName();
                visitedAssemblies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                componentHierarchy = new List<ComponentInfo>();
                processedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                int docType = swModel.GetType();
                if (docType == (int)swDocumentTypes_e.swDocASSEMBLY)
                {
                    if (swModel is IAssemblyDoc swAssembly)
                    {
                        BuildComponentHierarchy(swApp, swAssembly, 0);
                        ProcessByHierarchy(swApp, displaySettings);
                    }
                }
                else if (docType == (int)swDocumentTypes_e.swDocPART)
                {
                    ApplyDisplaySettings(swModel, displaySettings);
                }

                // 重新激活原始文档
                if (!(swApp.ActiveDoc is IModelDoc2 activeDoc &&
                      string.Equals(activeDoc.GetPathName(), originalOpenDocsPath, StringComparison.OrdinalIgnoreCase)))
                {
                    int err = 0;
                    swApp.ActivateDoc3(originalOpenDocsPath, false, 1, ref err);
                }

                MessageBox.Show("显示设置已应用到所有文档", "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message, "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Dictionary<swUserPreferenceToggle_e, bool> CreateDisplaySettings()
        {
            var settings = new Dictionary<swUserPreferenceToggle_e, bool>
            {
                [swUserPreferenceToggle_e.swDisplayPlanes] = false,             //基准面
                [swUserPreferenceToggle_e.swDisplayAxes] = false,               //轴
                [swUserPreferenceToggle_e.swDisplayCoordSystems] = false,       //坐标系
                [swUserPreferenceToggle_e.swDisplayOrigins] = false,            //原点
                [swUserPreferenceToggle_e.swDisplayCurves] = false,             //曲线
                [swUserPreferenceToggle_e.swGridDisplay] = false,               //网格
                [swUserPreferenceToggle_e.swDisplayCameras] = false,            //相机
                [swUserPreferenceToggle_e.swDisplayReferencePoints] = false,    //参考点
                [swUserPreferenceToggle_e.swDisplayWeldBead] = true,            //焊缝
                [swUserPreferenceToggle_e.swViewDispGlobalBBox] = false,        //全局边界框
                [swUserPreferenceToggle_e.swDisplayReferencePoints2] = false,   //参考点2
                [swUserPreferenceToggle_e.swDisplayDatumCoordSystems] = false,  //基准坐标系
                [swUserPreferenceToggle_e.swDisplayPartingLines] = false,       //分隔线
                [swUserPreferenceToggle_e.swDisplaySketches] = false,           //草图
                [swUserPreferenceToggle_e.swDisplaySketchPlanes] = false,       //草图基准面
                [swUserPreferenceToggle_e.swDisplayLights] = false,             //灯光
                [swUserPreferenceToggle_e.swDisplayLiveSections] = false,       //动态分隔线
                [swUserPreferenceToggle_e.swShowDimensionNames] = false,        //尺寸名称
                [swUserPreferenceToggle_e.swDisplayCenterOfMassSymbol] = false, //质心符号
                [swUserPreferenceToggle_e.swDisplayBendLines] = false,          //折弯线
                [swUserPreferenceToggle_e.swDisplayDecals] = false,             //装饰
                [swUserPreferenceToggle_e.swViewDisplayHideAllTypes] = false,   //隐藏所有类型
                [swUserPreferenceToggle_e.swDisplayAllAnnotations] = true,      //显示所有注释
                [swUserPreferenceToggle_e.swHideShowSketchDimensions] = true,   //显示草图尺寸
                [swUserPreferenceToggle_e.swDisplayCompAnnotations] = true,     //显示组件注释
                [swUserPreferenceToggle_e.swViewSketchRelations] = true,        //显示草图关系
            };
            return settings;
        }

        // 构建组件层级关系表
        private void BuildComponentHierarchy(ISldWorks swApp, IAssemblyDoc swAssembly, int level)
        {
            if (!(swAssembly is IModelDoc2 asmModel)) return;

            string compPath = asmModel.GetPathName();

            if (visitedAssemblies.Contains(compPath))
                return;
            visitedAssemblies.Add(compPath);

            componentHierarchy.Add(new ComponentInfo { Path = compPath, Level = level, Type = "ASSEMBLY" });

            object vComponents = swAssembly.GetComponents(false);
            if (vComponents == null) return;

            if (!(vComponents is object[] compArray)) return;

            for (int i = 0; i < compArray.Length; i++)
            {
                if (!(compArray[i] is IComponent2 swComp)) continue;

                bool isVirtual = false, isEnvelope = false;
                try { isVirtual = swComp.IsVirtual; } catch { }
                try { isEnvelope = swComp.IsEnvelope(); } catch { }
                if (isVirtual || isEnvelope || DocumentOps.IsSpeedPak(swComp))
                    continue;

                string childPath = swComp.GetPathName();
                if (string.IsNullOrEmpty(childPath) || !System.IO.File.Exists(childPath))
                    continue;

                string compName = swComp.Name;
                if (compName != null && compName.StartsWith("W_", StringComparison.OrdinalIgnoreCase))
                    continue;

                IModelDoc2 swChildModel = swComp.GetModelDoc2() as IModelDoc2;

                string compType;
                if (swChildModel == null)
                {
                    string ext = System.IO.Path.GetExtension(childPath).ToLower();
                    compType = (ext == ".sldasm") ? "ASSEMBLY" : "PART";
                }
                else
                {
                    int type = swChildModel.GetType();
                    compType = (type == (int)swDocumentTypes_e.swDocASSEMBLY) ? "ASSEMBLY" : "PART";
                }

                componentHierarchy.Add(new ComponentInfo { Path = childPath, Level = level + 1, Type = compType });

                if (compType == "ASSEMBLY")
                {
                    if (swChildModel is IAssemblyDoc childAssembly)
                    {
                        BuildComponentHierarchy(swApp, childAssembly, level + 1);
                    }
                    else if (DocumentOps.OpenDocumentSilent(swApp, childPath) is IAssemblyDoc childAsmSilent)
                    {
                        BuildComponentHierarchy(swApp, childAsmSilent, level + 1);
                    }
                }
            }
        }

        // 按层级从深到浅处理组件（应用显示设置）
        private void ProcessByHierarchy(ISldWorks swApp, Dictionary<swUserPreferenceToggle_e, bool> displaySettings)
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

                    IModelDoc2 swModel = DocumentOps.GetOpenedDocument(swApp, comp.Path)
                        ?? DocumentOps.OpenDocumentSilent(swApp, comp.Path);

                    if (swModel != null)
                    {
                        ApplyDisplaySettings(swModel, displaySettings);
                    }
                }
            }
        }

        // 应用显示设置到文档
        private void ApplyDisplaySettings(IModelDoc2 swModel, Dictionary<swUserPreferenceToggle_e, bool> displaySettings)
        {
            string modelPath = swModel.GetPathName();

            if (!(swApp.ActiveDoc is IModelDoc2 activeDoc &&
                  string.Equals(activeDoc.GetPathName(), modelPath, StringComparison.OrdinalIgnoreCase)))
            {
                int err = 0;
                swApp.ActivateDoc3(modelPath, false, 1, ref err);
            }

            foreach (var kvp in displaySettings)
            {
                try
                {
                    swModel.SetUserPreferenceToggle((int)kvp.Key, kvp.Value);
                }
                catch (Exception)
                {}
            }

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
    }
}
