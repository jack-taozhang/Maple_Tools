using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace Mtools
{
    /// <summary>
    /// SolidWorks 命令组图标生成（静态）
    /// 生成横向拼接的 BMP 图标列表文件，保存到 DLL 同目录 Icons 文件夹
    /// 图标0：模型更新  图标1：显示控制  图标2：赋二维码  图标3：图纸转档
    /// 新增功能图标：在 CreateIconList 中追加 DrawXxxIcon 调用并新增绘制方法
    /// </summary>
    internal static class IconGenerator
    {
        internal static string CreateIconList(int iconSize)
        {
            int count = 4;
            int totalWidth = iconSize * count;
            string iconsDir = Path.Combine(Path.GetDirectoryName(
                System.Reflection.Assembly.GetExecutingAssembly().Location), "Icons");
            Directory.CreateDirectory(iconsDir);
            string path = Path.Combine(iconsDir, "Mtools_" + iconSize + ".bmp");

            using (Bitmap bmp = new Bitmap(totalWidth, iconSize, PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                // 与 SwAddin.cs AddCommandMgr 顺序一致：模型更新→显示控制→赋二维码→图纸转档
                DrawUpdateModelsIcon(g, 0, 0, iconSize);
                DrawControlDisplayIcon(g, iconSize, 0, iconSize);
                DrawCreateQRCodeIcon(g, iconSize * 2, 0, iconSize);
                DrawExportDwgPdfIcon(g, iconSize * 3, 0, iconSize);

                bmp.Save(path, ImageFormat.Bmp);
            }
            return path;
        }

        // 圆角矩形路径
        private static GraphicsPath GetRoundedRectPath(int x, int y, int w, int h, float r)
        {
            GraphicsPath path = new GraphicsPath();
            path.AddArc(x, y, r * 2, r * 2, 180, 90);
            path.AddArc(x + w - r * 2, y, r * 2, r * 2, 270, 90);
            path.AddArc(x + w - r * 2, y + h - r * 2, r * 2, r * 2, 0, 90);
            path.AddArc(x, y + h - r * 2, r * 2, r * 2, 90, 90);
            path.CloseFigure();
            return path;
        }

        // 图标0：图纸转档（蓝色渐变 + 白色文档 + 折角 + 文字横线）
        private static void DrawExportDwgPdfIcon(Graphics g, int x, int y, int size)
        {
            float radius = size * 0.12f;
            using (GraphicsPath bgPath = GetRoundedRectPath(x, y, size, size, radius))
            using (LinearGradientBrush bgBrush = new LinearGradientBrush(
                new Rectangle(x, y, size, size),
                Color.FromArgb(59, 130, 246), Color.FromArgb(29, 78, 216), 90f))
            {
                g.FillPath(bgBrush, bgPath);
            }

            float docW = size * 0.46f;
            float docH = size * 0.58f;
            float docX = x + (size - docW) / 2f;
            float docY = y + (size - docH) / 2f;
            float fold = size * 0.15f;

            using (GraphicsPath docPath = new GraphicsPath())
            {
                docPath.AddLine(docX, docY, docX + docW - fold, docY);
                docPath.AddLine(docX + docW - fold, docY, docX + docW, docY + fold);
                docPath.AddLine(docX + docW, docY + fold, docX + docW, docY + docH);
                docPath.AddLine(docX + docW, docY + docH, docX, docY + docH);
                docPath.AddLine(docX, docY + docH, docX, docY);
                using (SolidBrush docBrush = new SolidBrush(Color.White))
                {
                    g.FillPath(docBrush, docPath);
                }
            }

            using (GraphicsPath foldPath = new GraphicsPath())
            {
                foldPath.AddLine(docX + docW - fold, docY, docX + docW, docY + fold);
                foldPath.AddLine(docX + docW, docY + fold, docX + docW - fold, docY + fold);
                foldPath.AddLine(docX + docW - fold, docY + fold, docX + docW - fold, docY);
                using (LinearGradientBrush foldBrush = new LinearGradientBrush(
                    new Rectangle((int)(docX + docW - fold), (int)docY, (int)fold, (int)fold),
                    Color.FromArgb(147, 197, 253), Color.FromArgb(59, 130, 246), 135f))
                {
                    g.FillPath(foldBrush, foldPath);
                }
            }

            float lineMargin = size * 0.07f;
            float lineStartX = docX + lineMargin;
            float lineEndX = docX + docW - lineMargin;
            float lineStartY = docY + fold + size * 0.06f;
            using (Pen linePen = new Pen(Color.FromArgb(59, 130, 246), size * 0.03f))
            {
                for (int i = 0; i < 3; i++)
                {
                    float ly = lineStartY + i * size * 0.085f;
                    g.DrawLine(linePen, lineStartX, ly, lineEndX, ly);
                }
            }
        }

        // 图标1：显示控制（绿色渐变 + 白色眼睛 + 瞳孔高光）
        private static void DrawControlDisplayIcon(Graphics g, int x, int y, int size)
        {
            float radius = size * 0.12f;
            using (GraphicsPath bgPath = GetRoundedRectPath(x, y, size, size, radius))
            using (LinearGradientBrush bgBrush = new LinearGradientBrush(
                new Rectangle(x, y, size, size),
                Color.FromArgb(34, 197, 94), Color.FromArgb(4, 120, 87), 90f))
            {
                g.FillPath(bgBrush, bgPath);
            }

            float cx = x + size / 2f;
            float cy = y + size / 2f;
            float eyeW = size * 0.58f;
            float eyeH = size * 0.34f;
            float pupilR = size * 0.11f;

            using (Pen eyePen = new Pen(Color.White, size * 0.055f))
            {
                g.DrawEllipse(eyePen, cx - eyeW / 2f, cy - eyeH / 2f, eyeW, eyeH);
            }

            using (SolidBrush pupilBrush = new SolidBrush(Color.White))
            {
                g.FillEllipse(pupilBrush, cx - pupilR, cy - pupilR, pupilR * 2f, pupilR * 2f);
            }

            float innerR = size * 0.05f;
            using (SolidBrush innerBrush = new SolidBrush(Color.FromArgb(4, 120, 87)))
            {
                g.FillEllipse(innerBrush, cx - innerR, cy - innerR, innerR * 2f, innerR * 2f);
            }

            float hlR = size * 0.025f;
            float hlOffset = size * 0.04f;
            using (SolidBrush hlBrush = new SolidBrush(Color.White))
            {
                g.FillEllipse(hlBrush, cx - hlOffset, cy - hlOffset, hlR * 2f, hlR * 2f);
            }
        }

        // 图标2：模型更新（橙色渐变 + 白色 Excel 表格 + 数据点）
        private static void DrawUpdateModelsIcon(Graphics g, int x, int y, int size)
        {
            float radius = size * 0.12f;
            using (GraphicsPath bgPath = GetRoundedRectPath(x, y, size, size, radius))
            using (LinearGradientBrush bgBrush = new LinearGradientBrush(
                new Rectangle(x, y, size, size),
                Color.FromArgb(249, 115, 22), Color.FromArgb(194, 65, 12), 90f))
            {
                g.FillPath(bgBrush, bgPath);
            }

            float tblW = size * 0.52f;
            float tblH = size * 0.5f;
            float tblX = x + (size - tblW) / 2f;
            float tblY = y + (size - tblH) / 2f;
            float lineW = size * 0.028f;

            using (SolidBrush tblBrush = new SolidBrush(Color.White))
            {
                g.FillRectangle(tblBrush, tblX, tblY, tblW, tblH);
            }

            float headerH = size * 0.12f;
            using (Pen orangePen = new Pen(Color.FromArgb(249, 115, 22), lineW))
            {
                g.DrawLine(orangePen, tblX, tblY + headerH, tblX + tblW, tblY + headerH);
                float colW = tblW / 3f;
                for (int i = 1; i < 3; i++)
                {
                    float cx2 = tblX + colW * i;
                    g.DrawLine(orangePen, cx2, tblY, cx2, tblY + tblH);
                }
                float rowH = (tblH - headerH) / 2f;
                for (int i = 1; i <= 2; i++)
                {
                    float cy2 = tblY + headerH + rowH * i;
                    g.DrawLine(orangePen, tblX, cy2, tblX + tblW, cy2);
                }
            }

            float dotR = size * 0.022f;
            using (SolidBrush dotBrush = new SolidBrush(Color.FromArgb(249, 115, 22)))
            {
                float rowH = (tblH - headerH) / 2f;
                float colW = tblW / 3f;
                for (int r = 0; r < 2; r++)
                {
                    for (int c = 0; c < 3; c++)
                    {
                        float dotCx = tblX + colW * c + colW / 2f;
                        float dotCy = tblY + headerH + rowH * r + rowH / 2f;
                        g.FillEllipse(dotBrush, dotCx - dotR, dotCy - dotR, dotR * 2f, dotR * 2f);
                    }
                }
            }
        }

        // 图标3：赋二维码（紫色渐变 + 白色 QR 码图案（三个定位方块 + 中间点阵））
        private static void DrawCreateQRCodeIcon(Graphics g, int x, int y, int size)
        {
            float radius = size * 0.12f;
            using (GraphicsPath bgPath = GetRoundedRectPath(x, y, size, size, radius))
            using (LinearGradientBrush bgBrush = new LinearGradientBrush(
                new Rectangle(x, y, size, size),
                Color.FromArgb(168, 85, 247), Color.FromArgb(109, 40, 217), 90f))
            {
                g.FillPath(bgBrush, bgPath);
            }

            // QR 码外框（白色正方形，内边距 18%）
            float pad = size * 0.18f;
            float qrX = x + pad;
            float qrY = y + pad;
            float qrSize = size - pad * 2;
            using (SolidBrush qrBg = new SolidBrush(Color.White))
            {
                g.FillRectangle(qrBg, qrX, qrY, qrSize, qrSize);
            }

            // 定位方块尺寸 = 3/7 的 QR 边长（标准 QR 定位角）
            float posSize = qrSize * 0.34f;
            float posInset = qrSize * 0.06f;

            using (Pen thickPen = new Pen(Color.FromArgb(76, 29, 149), qrSize * 0.06f))
            using (SolidBrush solidFill = new SolidBrush(Color.FromArgb(76, 29, 149)))
            {
                // 左上角定位方块
                DrawPositionMarker(g, thickPen, solidFill, qrX + posInset, qrY + posInset, posSize);
                // 右上角定位方块
                DrawPositionMarker(g, thickPen, solidFill,
                    qrX + qrSize - posInset - posSize, qrY + posInset, posSize);
                // 左下角定位方块
                DrawPositionMarker(g, thickPen, solidFill,
                    qrX + posInset, qrY + qrSize - posInset - posSize, posSize);
            }

            // 中间的小方块点阵（模拟编码数据区，填充紫黑色方块）
            using (SolidBrush dataFill = new SolidBrush(Color.FromArgb(91, 33, 182)))
            {
                float unit = qrSize * 0.065f;
                // 中间区域（避开三个定位方块及其保护区）：
                int unitX0 = (int)Math.Round(posSize / unit) + 2;  // 约占 5 个方格 + 保护区
                int unitXN = (int)Math.Round(qrSize / unit) - unitX0;  // 右侧也留保护区
                int unitY0 = unitX0;
                int unitYN = unitXN;

                // 用伪随机但稳定的图案（按坐标模）
                for (int iy = unitY0; iy < unitYN; iy++)
                {
                    for (int ix = unitX0; ix < unitXN; ix++)
                    {
                        // 避开右上和左下定位方块的保护区
                        if (ix > unitXN - (unitX0 + 2) && iy < unitY0 + (unitX0 + 2)) continue;
                        if (ix < unitX0 + (unitX0 + 2) && iy > unitYN - (unitY0 + 2)) continue;
                        // 确定性棋盘填充：约 50% 填满
                        bool filled = ((ix * 73856093) ^ (iy * 19349663)) % 3 == 0;
                        if (!filled) continue;
                        float bx = qrX + ix * unit;
                        float by = qrY + iy * unit;
                        g.FillRectangle(dataFill, bx, by, unit * 0.95f, unit * 0.95f);
                    }
                }
            }
        }

        // 绘制一个 QR 定位方块（外框 + 中间实心方块，3 层结构）
        private static void DrawPositionMarker(Graphics g, Pen outerPen, SolidBrush innerBrush,
            float x, float y, float size)
        {
            float outerW = outerPen.Width;
            // 外框（方形轮廓 = 外层粗边框）
            g.DrawRectangle(outerPen, x, y, size, size);
            // 中间实心方块（占中心 50% 大小，留出白色内环）
            float innerSize = size * 0.48f;
            float innerOffset = (size - innerSize) / 2f;
            g.FillRectangle(innerBrush, x + innerOffset, y + innerOffset, innerSize, innerSize);
            // 中间再加一个内边框（让视觉层次更清晰）
            using (Pen midPen = new Pen(innerBrush.Color, outerW * 0.7f))
            {
                float border = size * 0.22f;
                g.DrawRectangle(midPen,
                    x + border, y + border, size - border * 2, size - border * 2);
            }
        }
    }
}
