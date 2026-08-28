using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    [ComVisible(true)]
    [Guid("552ACFC2-6261-4D10-A3D9-0927A255842B")]
    [ProgId("Mtools.SwAddin")]
    // AutoDispatch 只暴露 IDispatch（后期绑定），避免 tlbexp 尝试解析
    // SolidWorks.Interop 接口签名触发 "找不到元素" 错误；
    // SW 通过 SetAddinCallbackInfo + 字符串名反射方式调用回调方法，
    // 走 IDispatch::GetIDsOfNames/Invoke 即可，不需要 dual/vtable 接口。
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public partial class SwAddin : SolidWorks.Interop.swpublished.SwAddin
    {
        private ISldWorks swApp;
        private int addinID;
        private ICommandManager cmdMgr;
        private ICommandGroup cmdGroup;

        private const int CMD_GROUP_ID = 875231;

        // ========== 共享字段（显示控制 & 模型更新功能使用）==========
        internal string originalOpenDocsPath;
        internal HashSet<string> visitedAssemblies;
        internal List<ComponentInfo> componentHierarchy;
        internal HashSet<string> processedFiles;
        internal HashSet<string> originallyOpenedDocs;

        // 内部纯 .NET 数据结构，禁止封送到 COM 类型库，防止 tlbexp 找不到类型
        [ComVisible(false)]
        internal class ComponentInfo
        {
            public string Path;
            public int Level;
            public string Type;
        }

        // 统一日志（静态，供本类及工具类调用）
        // 日志已关闭（no-op），如需排错可恢复实现
        internal static void Log()
        {
            // no-op
        }

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            swApp = ThisSW as ISldWorks;
            if (swApp == null)
            {
                return false;
            }

            addinID = Cookie;
            swApp.SetAddinCallbackInfo(0, this, addinID);

            cmdMgr = swApp.GetCommandManager(addinID);
            if (cmdMgr == null)
            {
                return false;
            }

            bool result = AddCommandMgr();
            return result;
        }

        public bool DisconnectFromSW()
        {
            RemoveCommandMgr();
            swApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            return true;
        }

        // ========== 命令管理 ==========
        // 新增功能步骤：
        //   1. 在 Commands 文件夹新建 SwAddin.Xxx.cs（partial）实现回调方法
        //   2. 在下方 AddCommandMgr 中添加 AddCommandItem2 注册项
        //   3. 在下方 cmdIDs/textTypes 数组中追加对应命令 ID
        //   4. 在 UI/IconGenerator.cs 中追加图标绘制方法
        private bool AddCommandMgr()
        {
            int errors = 0;

            // 先移除旧的命令组（防止缓存冲突）
            try
            {
                cmdMgr.RemoveCommandGroup(875231);
                cmdMgr.RemoveCommandGroup(147852);
            }
            catch { }

            cmdGroup = cmdMgr.CreateCommandGroup2(
                CMD_GROUP_ID,
                "Mtools",
                "Mtools 插件",
                "Mtools 插件命令",
                -1,
                true,
                ref errors);

            if (cmdGroup == null)
            {
                return false;
            }

            // 生成并设置命令按钮图标
            try
            {
                string largeIconPath = IconGenerator.CreateIconList(24);
                string smallIconPath = IconGenerator.CreateIconList(16);
                cmdGroup.LargeIconList = largeIconPath;
                cmdGroup.SmallIconList = smallIconPath;
            }
            catch (Exception)
            {}

            // 按钮顺序：
            //  组1（模型操作）：1.模型更新  2.显示控制
            //  组2（图纸输出）：3.赋二维码  4.图纸转档
            int cmdIndex1 = cmdGroup.AddCommandItem2(
                "UpdateModels",
                -1,
                "Excel参数更新模型方程式",
                "模型更新",
                0,                          // 图标索引 0
                "UpdateModelsFromExcel",     // 回调方法名
                "",
                1,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            if (cmdIndex1 < 0)
            {
                return false;
            }

            int cmdIndex2 = cmdGroup.AddCommandItem2(
                "ControlDisplay",
                -1,
                "控制SW模型中类型是否显示",
                "显示设置",
                1,                          // 图标索引 1
                "ControlSolidWorksDisplay",  // 回调方法名
                "",
                2,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            if (cmdIndex2 < 0)
            {
                return false;
            }

            int cmdIndex3 = cmdGroup.AddCommandItem2(
                "GenerateQR",
                -1,
                "生成二维码并插入工程图",
                "赋二维码",
                2,                          // 图标索引 2
                "GenerateQR",                // 回调方法名
                "EnableGenerateQR",          // 启用条件：仅工程图
                3,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            if (cmdIndex3 < 0)
            {
                return false;
            }

            int cmdIndex4 = cmdGroup.AddCommandItem2(
                "ExportDwgPdf",
                -1,
                "导出DWG和PDF",
                "图纸转档",
                3,                          // 图标索引 3
                "ExportDwgPdf",              // 回调方法名
                "EnableExportDwgPdf",        // 启用条件：仅工程图
                4,
                (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem));

            if (cmdIndex4 < 0)
            {
                return false;
            }

            cmdGroup.HasMenu = true;
            cmdGroup.HasToolbar = true;  // 必须为 true，选项卡按钮的命令 ID 来自命令组
            cmdGroup.Activate();

            // 在 CommandManager 中创建选项卡（Part=1, Assembly=2, Drawing=3）
            int[] docTypes = {
                (int)swDocumentTypes_e.swDocPART,
                (int)swDocumentTypes_e.swDocASSEMBLY,
                (int)swDocumentTypes_e.swDocDRAWING
            };

            // 顺序与 AddCommandItem2 注册顺序一致：更新→显示→QR→转档
            int cmdId1 = cmdGroup.get_CommandID(0);  // UpdateModelsFromExcel
            int cmdId2 = cmdGroup.get_CommandID(1);  // ControlSolidWorksDisplay
            int cmdId3 = cmdGroup.get_CommandID(2);  // GenerateQR
            int cmdId4 = cmdGroup.get_CommandID(3);  // ExportDwgPdf
            int[] cmdIDs = { cmdId1, cmdId2, cmdId3, cmdId4 };
            int[] textTypes = {
                (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow,
                (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow
            };

            foreach (int docType in docTypes)
            {
                try
                {
                    if (cmdMgr.GetCommandTab(docType, "Mtools") is CommandTab existingTab)
                    {
                        cmdMgr.RemoveCommandTab(existingTab);
                    }

                    if (!(cmdMgr.AddCommandTab(docType, "Mtools") is CommandTab cmdTab))
                    {
                        continue;
                    }

                    if (!(cmdTab.AddCommandTabBox() is ICommandTabBox cmdBox))
                    {
                        continue;
                    }

                    cmdBox.AddCommands(cmdIDs, textTypes);

                    // 在"赋二维码（组2第一个）"【前面】加分隔线 = 正好在显示控制(组1最后一个)之后
                    // 视觉区分"模型操作组 / 图纸输出组"：
                    // [0]模型更新  [1]显示控制  ▏(分隔此处)  [2]赋二维码  [3]图纸转档
                    try
                    {
                        // AddSeparator(tabBox, CommandID) 语义:在【指定 CommandID 对应按钮的前面】插入分隔
                        // 所以使用 cmdId3 = GenerateQR（赋二维码）的 ID → 分隔就在赋二维码前面。
                        if (cmdBox is CommandTabBox tabBoxClass)
                        {
                            cmdTab.AddSeparator(tabBoxClass, cmdId3);
                        }
                    }
                    catch (Exception)
                    {}
                }
                catch (Exception)
                {}
            }
            return true;
        }

        private void RemoveCommandMgr()
        {
            if (cmdMgr != null)
            {
                int[] docTypes = {
                    (int)swDocumentTypes_e.swDocPART,
                    (int)swDocumentTypes_e.swDocASSEMBLY,
                    (int)swDocumentTypes_e.swDocDRAWING
                };
                foreach (int docType in docTypes)
                {
                    try
                    {
                        if (cmdMgr.GetCommandTab(docType, "Mtools") is CommandTab tab)
                        {
                            cmdMgr.RemoveCommandTab(tab);
                        }
                    }
                    catch (Exception)
                    {}
                }

                if (cmdGroup != null)
                {
                    try
                    {
                        cmdMgr.RemoveCommandGroup(CMD_GROUP_ID);
                    }
                    catch (Exception)
                    {}
                    cmdGroup = null;
                }
            }
        }

        // ========== COM 注册 ==========
        // 注意：SW 启动时从 HKLM\SOFTWARE\SolidWorks\AddIns\{GUID} 枚举 addin，
        // (default) 空名值=1 表示启用自动加载。HKCU 同步注册作为用户级覆盖。
        // regasm 以 admin 运行（VS admin 构建）时同时写 HKLM + HKCU；
        // 非 admin 时只写 HKCU（SW 仍可从 HKLM 旧项加载，需手工迁移）。
        [ComRegisterFunction]
        private static void RegisterFunction(Type t)
        {
            string guidString = t.GUID.ToString("B").ToUpper();
            string swAddInPath = @"SOFTWARE\SolidWorks\AddIns\" + guidString;

            // HKLM 机器级注册（SW 启动时枚举此键，必须）
            try
            {
                using (RegistryKey key = Registry.LocalMachine.CreateSubKey(swAddInPath))
                {
                    key.SetValue(null, 1, RegistryValueKind.DWord);
                    key.SetValue("Title", "Mtools");
                    key.SetValue("Description", "Mtools 插件");
                }
            }
            catch (Exception)
            {}

            // HKCU 用户级覆盖（无需 admin，备份）
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(
                @"Software\SolidWorks\AddIns\" + guidString))
            {
                key.SetValue(null, 1, RegistryValueKind.DWord);
                key.SetValue("Title", "Mtools");
                key.SetValue("Description", "Mtools 插件");
            }
        }

        [ComUnregisterFunction]
        private static void UnregisterFunction(Type t)
        {
            string guidString = t.GUID.ToString("B").ToUpper();
            string swAddInPath = @"SOFTWARE\SolidWorks\AddIns\" + guidString;

            try
            {
                Registry.LocalMachine.DeleteSubKey(swAddInPath, false);
            }
            catch (Exception)
            {}
            Registry.CurrentUser.DeleteSubKey(
                @"Software\SolidWorks\AddIns\" + guidString, false);
        }
    }
}
