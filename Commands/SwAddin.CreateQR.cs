using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using QRCoder;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace Mtools
{
    /// <summary>
    /// 赋二维码功能：读取工程图图号中 TS 开头数据，生成 QR 码并插入到工程图右上角，
    /// 在二维码下方添加文本显示。
    /// </summary>
    public partial class SwAddin
    {
        // ====================================================================
        // 命令回调（在 SwAddin.cs AddCommandMgr 中通过字符串名注册）
        // ====================================================================

        /// <summary>
        /// 生成二维码按钮的启用条件：当前文档为工程图时可用
        /// </summary>
        public int EnableGenerateQR()
        {
            if (swApp.ActiveDoc == null) return 0;
            ModelDoc2 modDoc = (ModelDoc2)swApp.ActiveDoc;
            return (modDoc.GetType() == (int)swDocumentTypes_e.swDocDRAWING) ? 1 : 0;
        }

        /// <summary>
        /// 命令回调：赋二维码（从工程图图号提取 TS 开头内容，生成 QR 并插入图纸）
        /// </summary>
        public void GenerateQR()
        {

            try
            {
                // 1. 校验当前文档是否为工程图 (SLDDRW)
                ModelDoc2 modDoc = (ModelDoc2)swApp.ActiveDoc;
                if (modDoc == null)
                {
                    MessageBox.Show("请先打开工程图文档。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (modDoc.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
                {
                    MessageBox.Show("当前文档不是工程图，无法生成二维码。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 校验文件扩展名是否为 SLDDRW
                string filePath = modDoc.GetPathName();
                if (string.IsNullOrEmpty(filePath))
                {
                    MessageBox.Show("当前文档未保存，请先保存为 SLDDRW 文件。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (!filePath.EndsWith(".slddrw", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("当前文档不是 SLDDRW 文件。", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DrawingDoc drawDoc = (DrawingDoc)modDoc;

                // 赋二维码前：图面布满显示（整屏 / 适应窗口）
                ZoomDrawingToFit(modDoc);

                // 2. 读取图纸模型属性信息，查找图号
                string drawingNumber = GetDrawingNumber(modDoc);

                // 3. 若图号为空则提示
                if (string.IsNullOrWhiteSpace(drawingNumber))
                {
                    MessageBox.Show("无图号", "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 4. 截取图号数据 TS 开头至结尾
                string qrContent = ExtractTsData(drawingNumber);
                if (string.IsNullOrEmpty(qrContent))
                {
                    MessageBox.Show("图号中未找到 TS 开头的数据。\n图号: " + drawingNumber, "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 5. 生成 QR 码图片
                string qrImagePath;
                try
                {
                    qrImagePath = GenerateQRCodeImage(qrContent);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("二维码生成失败：" + ex.Message, "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 6. 将二维码插入到图纸
                try
                {
                    InsertQRCodeToDrawing(drawDoc, qrImagePath, qrContent);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("二维码插入失败：" + ex.Message, "Mtools",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    try { if (File.Exists(qrImagePath)) File.Delete(qrImagePath); } catch { }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message, "Mtools",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 图面布满显示：将当前工程图缩放至充满整个图形区域（等同"整屏显示 / 适应窗口"）。
        /// 仅影响屏幕显示，不改变图纸几何与导出比例。
        /// </summary>
        private void ZoomDrawingToFit(IModelDoc2 modDoc)
        {
            if (modDoc == null) return;
            try
            {
                // ViewZoomtofit2 将模型/工程图缩放至适配图形区
                modDoc.ViewZoomtofit2();
            }
            catch (Exception) { }
        }

        // ====================================================================
        // 辅助方法
        // ====================================================================

        /// <summary>
        /// 读取工程图引用模型的自定义属性中的图号（图号 / DrawingNumber / PartNumber / 零件编号）
        /// </summary>
        private string GetDrawingNumber(ModelDoc2 modDoc)
        {
            string[] propertyNames = { "图号", "DrawingNumber", "PartNumber", "零件编号" };

            if (!(modDoc is DrawingDoc drawDoc)) return "";

            // 获取第一个工程视图（GetFirstView 返回图纸本身，NextView 才是实际视图）
            SolidWorks.Interop.sldworks.View firstView = (SolidWorks.Interop.sldworks.View)drawDoc.GetFirstView();
            if (firstView != null)
            {
                firstView = (SolidWorks.Interop.sldworks.View)firstView.GetNextView();
            }
            if (firstView == null) return "";

            // 获取视图引用的模型
            ModelDoc2 refModel = null;
            try
            {
                refModel = firstView.ReferencedDocument;
            }
            catch { }

            // ReferencedDocument 为空时尝试通过模型名打开
            if (refModel == null)
            {
                try
                {
                    string refModelName = firstView.GetReferencedModelName();
                    if (!string.IsNullOrEmpty(refModelName))
                    {
                        int err = 0, warn = 0;
                        refModel = swApp.OpenDoc6(refModelName, (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
                        if (refModel == null)
                        {
                            refModel = swApp.OpenDoc6(refModelName, (int)swDocumentTypes_e.swDocASSEMBLY,
                                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
                        }
                    }
                }
                catch { }
            }
            if (refModel == null) return "";

            // 遍历配置读取自定义属性
            object[] configs = (object[])refModel.GetConfigurationNames();
            if (configs != null)
            {
                foreach (object configObj in configs)
                {
                    string configName = (string)configObj;
                    CustomPropertyManager cpm = refModel.Extension.CustomPropertyManager[configName];
                    if (cpm == null) continue;
                    foreach (string propName in propertyNames)
                    {
                        cpm.Get2(propName, out string valOut, out _);
                        if (!string.IsNullOrWhiteSpace(valOut))
                            return valOut.Trim();
                    }
                }
            }

            // 兜底：读取模型自身非配置特定属性
            foreach (string propName in propertyNames)
            {
                string value = "";
                try { value = refModel.CustomInfo2["", propName]; } catch { }
                if (!string.IsNullOrWhiteSpace(value))
                    return value.Trim();
            }

            return "";
        }

        /// <summary>
        /// 截取图号中从 TS 开头至结尾的数据
        /// </summary>
        private string ExtractTsData(string drawingNumber)
        {
            int tsIndex = drawingNumber.IndexOf("TS", StringComparison.Ordinal);
            if (tsIndex < 0) return "";
            return drawingNumber.Substring(tsIndex);
        }

        /// <summary>
        /// 使用 QRCoder 生成二维码 BMP 图片：
        /// - 黑模块 + 纯白底（不透明）
        /// - 边缘额外增加 2 像素白边框
        /// - 以 BMP 格式保存（无 alpha 通道），插入时白底为不透明白。
        /// </summary>
        private string GenerateQRCodeImage(string content)
        {
            string tempPath = Path.Combine(Path.GetTempPath(),
                "Mtools_QR_" + Guid.NewGuid().ToString("N") + ".bmp");

            using (QRCodeGenerator qrGenerator = new QRCodeGenerator())
            using (QRCodeData qrCodeData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q))
            using (QRCode qrCode = new QRCode(qrCodeData))
            // 第2参数 darkColor=黑，第3参数 lightColor=白：直接生成"黑模块 + 白底"位图；
            // drawQuietZones=false：不保留 QRCoder 自带的留白边（约 4 模块白边），
            // 白色仅由下方 2 像素边框提供，避免二维码四周白区过大。
            using (Bitmap rawBitmap = qrCode.GetGraphic(20, Color.Black, Color.White, false))
            {
                int border = 2; // 边缘白边框像素

                // 目标画布 = 原图 + 四周各 border 像素白边，24bpp RGB（无 alpha 通道）
                int w = rawBitmap.Width + border * 2;
                int h = rawBitmap.Height + border * 2;
                using (Bitmap bmpOut = new Bitmap(w, h, PixelFormat.Format24bppRgb))
                {
                    using (Graphics g = Graphics.FromImage(bmpOut))
                    {
                        // 先整体填白，再把二维码绘制到内缩 border 像素的位置 → 四周留出白边框
                        g.Clear(Color.White);
                        g.DrawImage(rawBitmap, border, border, rawBitmap.Width, rawBitmap.Height);
                    }

                    // 保存为 BMP
                    bmpOut.Save(tempPath, ImageFormat.Bmp);
                }
            }

            return tempPath;
        }

        // 二维码草图图片特征名标记：插入时写入，删除时据此可靠识别（无需坐标探针）
        private const string QR_FEATURE_MARKER = "MtoolsQRCode";

        /// <summary>
        /// 【按名称选中删除】—— 唯一保留的删除路径（官方论坛验证 + 用户实测有效）：
        /// 插入时用 SketchPicture.GetFeature().Name 统一命名（QR_FEATURE_MARKER），
        /// 删除时用 SelectByID2(名称, "SKETCHBITMAP", 0,0,0,...) 按名称精确选中后删除。
        /// 全程仅按名称查找，选不中立即退出——无坐标扫描、无全页遍历、无特征树遍历，
        /// 仅 2~3 次选择调用，速度最快。
        /// 同名可能累积多张（历史删除失败留下的），循环删到选不中为止。
        /// </summary>
        private int DeletePicturesByName(ModelDoc2 modDoc)
        {
            if (modDoc == null) return 0;

            int deleted = 0;
            for (int i = 0; i < 10; i++)
            {
                try
                {
                    modDoc.ClearSelection2(true);
                    if (!modDoc.Extension.SelectByID2(QR_FEATURE_MARKER, "SKETCHBITMAP", 0, 0, 0, false, 0, null, 0))
                        break;
                    if (modDoc.ISelectionManager.GetSelectedObjectCount2(-1) != 1)
                    { modDoc.ClearSelection2(true); break; }

                    modDoc.Extension.DeleteSelection2(1);
                    modDoc.ClearSelection2(true);
                    deleted++;
                }
                catch (Exception) { modDoc.ClearSelection2(true); break; }
            }
            return deleted;
        }

        /// <summary>
        /// 将二维码图片插入到当前工程图左下角（按图纸比例自适应），并在下方添加文本。
        /// </summary>
        private void InsertQRCodeToDrawing(DrawingDoc drawDoc, string imagePath, string qrContent)
        {
            ModelDoc2 modDoc = (ModelDoc2)drawDoc;

            // 获取当前图纸
            if (!(drawDoc.GetCurrentSheet() is Sheet currentSheet))
                throw new Exception("无法获取当前图纸。");

            // 激活当前图纸空间（非视图空间）
            string sheetName = currentSheet.GetName();
            drawDoc.ActivateSheet(sheetName);
            drawDoc.ActivateView("");
            modDoc.ClearSelection2(true);

            // ===== 读取图纸比例（用于计算二维码尺寸与位置）=====
            double scale1 = 1, scale2 = 1;
            object props = currentSheet.GetProperties2();
            if (props is double[] propArr && propArr.Length >= 7)
            {
                scale1 = propArr[2];
                scale2 = propArr[3];
            }

            // 比例系数 = 比例分母 / 比例分子（例 1:100 → 100，2:1 → 0.5）
            double scaleFactor = (scale1 != 0) ? (scale2 / scale1) : 1.0;

            const double QR_SIZE_M = 0.030;
            const double OFFSET_LEFT_M = 0.007;
            const double QR_POS_Y_M = 0.010;                // 二维码插入点 Y = 10mm
            const double TEXT_TARGET_X_M = OFFSET_LEFT_M;   // 文本左边缘 X = 7mm（与二维码插入点左对齐）
            const double TEXT_TARGET_TOP_Y_M = 0.010;       // 文本插入点 Y = 10mm（左上角对齐到插入点）
            const double TEXT_HEIGHT_M = 0.0035;            // 文本字体大小 = 3.5mm

            // 二维码图片尺寸和位置（与图纸比例相关）
            double qrSize = QR_SIZE_M * scaleFactor;
            double posX   = OFFSET_LEFT_M   * scaleFactor;
            double posY   = QR_POS_Y_M * scaleFactor;   // 插入点 Y = 10mm

            // ===== 删除上次插入的二维码图片和文字 =====
            // 1) Picture 删除：按名称选中删除（唯一保留路径，最快）。
            //    插入时已把图片特征命名为 QR_FEATURE_MARKER，
            //    用 SelectByID2(名称, "SKETCHBITMAP") 按名称精确选中后删除，无任何扫描。
            // 2) Note 删除：遍历所有 Note，找 TS 开头的删除
            try
            {
                // 1) 删除旧二维码图片（按名称，选不中立即退出）
                try { DeletePicturesByName(modDoc); }
                catch (Exception) { }

                // 2) 用反射调用 GetAnnotations 获取所有 Annotation，找 TS 开头的 Note 删除
                try
                {
                    object annosResult = null;
                    try
                    {
                        // 反射调用 IModelDocExtension.GetAnnotations()
                        annosResult = modDoc.Extension.GetType().InvokeMember(
                            "GetAnnotations",
                            BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                            null, modDoc.Extension, null);
                    }
                    catch (Exception) {}

                    if (annosResult is object[] annos && annos.Length > 0)
                    {
                        foreach (object annoObj in annos)
                        {
                            if (!(annoObj is Note note)) continue;

                            string text = note.GetText();
                            if (string.IsNullOrEmpty(text) || !text.StartsWith("TS")) continue;

                            if (!(note.GetAnnotation() is Annotation anno)) continue;

                            modDoc.ClearSelection2(true);
                            anno.Select3(false, null);
                            int selCount = modDoc.ISelectionManager.GetSelectedObjectCount2(-1);
                            if (selCount == 1)
                            {
                                modDoc.Extension.DeleteSelection2(1);
                                modDoc.ClearSelection2(true);
                            }
                        }
                    }
                    else
                    {
                        // GetAnnotations 返回 null，用 SelectByID2(type="NOTE") 选中删除
                        double noteX = TEXT_TARGET_X_M;
                        double noteY = TEXT_TARGET_TOP_Y_M;
                        for (int i = 0; i < 5; i++)
                        {
                            modDoc.ClearSelection2(true);
                            bool sel = modDoc.Extension.SelectByID2(
                                "", "NOTE", noteX, noteY, 0, false, 0, null, 0);
                            if (!sel) break;

                            int selCount = modDoc.ISelectionManager.GetSelectedObjectCount2(-1);
                            int objType = (selCount > 0) ? modDoc.ISelectionManager.GetSelectedObjectType3(1, -1) : -1;

                            // 获取选中对象，用反射获取文本，只有文本以 "TS" 开头才删除
                            if (selCount == 1)
                            {
                                object selObj = modDoc.ISelectionManager.GetSelectedObject6(1, -1);
                                string text = "";
                                try
                                {
                                    text = selObj.GetType().InvokeMember("GetText",
                                        BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                                        null, selObj, null).ToString();
                                }
                                catch (Exception) {}

                                if (!string.IsNullOrEmpty(text) && text.StartsWith("TS"))
                                {
                                    modDoc.Extension.DeleteSelection2(1);
                                    modDoc.ClearSelection2(true);
                                }
                                else
                                {
                                    modDoc.ClearSelection2(true);
                                    break;
                                }
                            }
                            else
                            {
                                modDoc.ClearSelection2(true);
                                break;
                            }
                        }
                    }
                }
                catch (Exception) {}
            }
            catch (Exception) {}

            modDoc.ClearSelection2(true);

            // 文本"纸面目标坐标"（常量，与比例无关）：
            //   X = 7mm（与二维码插入点左对齐），Y = 10mm
            //   CreateText2 的坐标即文本左上角，配合左对齐实现"左上对齐到插入点"
            double api_setPositionX = TEXT_TARGET_X_M;
            double api_setPositionY = TEXT_TARGET_TOP_Y_M;

            // 插入草图图片（在当前图纸空间）
            if (!(modDoc.SketchManager.InsertSketchPicture(imagePath) is SketchPicture sketchPic))
                throw new Exception("图片插入返回 null，请检查文件格式和权限。");

            sketchPic.SetOrigin(posX, posY);
            sketchPic.SetSize(qrSize, qrSize, false);
            sketchPic.Angle = 0;

            // 写入特征名标记，供下次插入时精确识别并删除（覆盖"旧二维码无法删除"问题）
            try
            {
                if (sketchPic.GetFeature() is IFeature qrFeature)
                {
                    qrFeature.Name = QR_FEATURE_MARKER;
                }
            }
            catch (Exception) { }

            // 【关键】设置草图图片透明度模式
            // 二维码已改为"白底 BMP"（无 alpha 通道），故关闭透明度，使白底为不透明白。
            // ISketchPicture.SetTransparency 签名（SW 2024+）：
            //   useTransparencyFromFile (int)    : 0 = 不使用文件透明通道（BMP 无 alpha）
            //   transparentColor        (double): 当 useTransparencyFromFile=0 时指定 RGB 透明色（0 表示无）
            //   useMatchingTolerance    (int)   : 1 = 启用颜色容差, 0 = 不启用
            //   matchingTolerance       (double): 颜色容差（0.0 ~ 1.0）
            // 这里 useTransparencyFromFile=0 → BMP 白底整体不透明显示，二维码为白底黑模块。
            try
            {
                sketchPic.SetTransparency(0, 0.0, 0, 0.0);
            }
            catch (Exception) {
                // 兼容性兜底：部分 SW 版本 API 签名或调用顺序不同，不影响主流程
            }

            // 按 ESC 关闭草图图片属性页
            try { swApp.RunCommand(-2, ""); } catch { }

            // 退出草图（true = 保持在图纸空间而非草图）
            modDoc.SketchManager.InsertSketch(true);
            modDoc.ClearSelection2(true);

            // 再次激活：保证后续 InsertNote2 在图纸空间而不是某个视图或草图平面上下文
            drawDoc.ActivateSheet(sheetName);
            drawDoc.ActivateView("");
            modDoc.ClearSelection2(true);

            double createText2_UpperLeftX = api_setPositionX;
            double createText2_UpperLeftY = api_setPositionY;
            object noteObj;
            try
            {
                noteObj = drawDoc.CreateText2(qrContent,
                    createText2_UpperLeftX, createText2_UpperLeftY, 0,
                    TEXT_HEIGHT_M, 0);
            }
            catch (Exception) {
                noteObj = null;
            }

            Note textNote;
            if (noteObj is Note createdNote)
            {
                textNote = createdNote;
            }
            else
            {
                // 回退 1：CreateText 同样给 ÷factor 的左上角坐标
                try
                {
                    drawDoc.CreateText(qrContent,
                        createText2_UpperLeftX, createText2_UpperLeftY, 0,
                        TEXT_HEIGHT_M, 0);
                }
                catch (Exception) {}

                // 回退 2：InsertNote + 强制定位
                modDoc.ClearSelection2(true);
                if (!(modDoc.InsertNote(qrContent) is Note insertedNote))
                {
                    throw new Exception("文本插入失败（CreateText2/InsertNote 均返回 null）。");
                }
                textNote = insertedNote;
            }

            // 左对齐：文本从插入点向右延伸（配合 CreateText2 左上角坐标 = 左上对齐到插入点）
            try { textNote.SetTextJustification((int)swTextJustification_e.swTextJustificationLeft); }
            catch (Exception) {}
            try { textNote.Angle = 0; } catch { }

            if (!(textNote.GetAnnotation() is Annotation textAnno))
            {
                throw new Exception("文本插入后获取 Annotation 失败。");
            }

            // 强制无引线（在 SetPosition2 之前设置，避免引线把位置拉回）
            try
            {
                textAnno.SetLeader3(0, 0, false, false, false, false);
            }
            catch (Exception) {}

            // 最终精确定位：
            //   X = 纸面 7mm（与二维码插入点左对齐，文本向右延伸）
            //   Y = 纸面 10mm
            textAnno.SetPosition2(api_setPositionX, api_setPositionY, 0);
            modDoc.ClearSelection2(true);

            // Note 删除依赖文本"TS"开头标记，不需要存储 PID
        }
    }
}
