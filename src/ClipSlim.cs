using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Globalization;
using System.IO.Pipes;
using System.Security.AccessControl;
using System.Security.Principal;

// ClipSlim: standalone Windows clipboard image compressor. MIT licensed.
internal static class Program
{
    [DllImport("user32.dll")] private static extern bool ShowWindowAsync(IntPtr handle, int command);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr handle);
    private static Mutex singleInstance;
    [STAThread]
    private static int Main(string[] args)
    {
        if (args.Length > 0 && args[0] == "--self-test") return Tests.Run(args.Length > 1 ? args[1] : ".");
        if (args.Length > 1 && args[0] == "--wps-test") return WpsBridge.Test(args[1]);
        if (args.Length > 1 && args[0] == "--layout-test")
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (MainWindow window = new MainWindow()) return window.LayoutTest(args[1]);
        }
        if (args.Length > 1 && args[0] == "--integration-test")
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            using (MainWindow window = new MainWindow()) return window.IntegrationTest(args[1]);
        }
        if (args.Length > 1 && args[0] == "--verify-clipboard")
        {
            try
            {
                using (Image image = Clipboard.GetImage())
                {
                    if (image == null || image.Width != 640 || image.Height != 360) throw new Exception("Clipboard image missing or wrong size");
                    File.WriteAllText(args[1], "PASS: persisted clipboard image readable by a separate process (640x360).", Encoding.UTF8);
                }
                return 0;
            }
            catch (Exception ex) { File.WriteAllText(args[1], "FAIL: " + ex, Encoding.UTF8); return 1; }
        }
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        bool created;
        singleInstance = new Mutex(false, "Local\\ClipSlim.Collage.v171." + WindowsIdentity.GetCurrent().User.Value, out created);
        if (!created)
        {
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("ClipSlim-v1.7.1"))
            {
                using (process) if (process.MainWindowHandle != IntPtr.Zero) { ShowWindowAsync(process.MainWindowHandle, 9); SetForegroundWindow(process.MainWindowHandle); break; }
            }
            singleInstance.Dispose();
            return 0;
        }
        Application.Run(new MainWindow());
        singleInstance.Dispose();
        return 0;
    }
}

internal sealed class Result : IDisposable
{
    public byte[] Bytes;
    public Bitmap Preview;
    public string Extension;
    public void Dispose() { if (Preview != null) Preview.Dispose(); }
}

internal sealed class WpsTarget : IDisposable
{
    public object App, Book, Sheet, Cell, Range;
    public string Label;
    public double Left, Top, Width, Height;
    public void Dispose()
    {
        if (!Object.ReferenceEquals(Cell, Range)) WpsBridge.Release(Range);
        WpsBridge.Release(Cell); WpsBridge.Release(Sheet); WpsBridge.Release(Book); WpsBridge.Release(App);
    }
}

internal static class WpsBridge
{
    public static Action<string> Trace;
    public static void Release(object value) { if (value != null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value); }
    public static object Connect()
    {
        try { return Marshal.GetActiveObject("ket.Application"); }
        catch (COMException) { throw new InvalidOperationException("未连接到 WPS 表格。请先打开桌面版 WPS 工作簿，并选中目标单元格。"); }
    }
    public static WpsTarget ReadTarget()
    {
        WpsTarget t = new WpsTarget();
        try
        {
            t.App = Connect(); dynamic app = t.App;
            t.Book = app.ActiveWorkbook;
            if (t.Book == null) throw new InvalidOperationException("WPS 尚未打开工作簿。");
            t.Sheet = app.ActiveSheet;
            t.Cell = app.ActiveCell;
            if (t.Cell == null) throw new InvalidOperationException("请先在 WPS 中选中一个单元格。");
            dynamic cell = t.Cell;
            t.Range = Convert.ToBoolean(cell.MergeCells) ? (object)cell.MergeArea : t.Cell;
            dynamic range = t.Range;
            t.Left = Convert.ToDouble(range.Left); t.Top = Convert.ToDouble(range.Top);
            t.Width = Convert.ToDouble(range.Width); t.Height = Convert.ToDouble(range.Height);
            t.Label = Convert.ToString(((dynamic)t.Book).Name) + " / " + Convert.ToString(((dynamic)t.Sheet).Name) + "!" + Convert.ToString(range.Address);
            return t;
        }
        catch { t.Dispose(); throw; }
    }
    public static object Insert(object sheetObject, object rangeObject, string file, int imageWidth, int imageHeight)
    {
        dynamic sheet = sheetObject, range = rangeObject;
        if (Convert.ToBoolean(sheet.ProtectDrawingObjects)) throw new InvalidOperationException("当前工作表的图形受保护，无法插入图片。");
        double cellWidth = Convert.ToDouble(range.Width), cellHeight = Convert.ToDouble(range.Height);
        if (cellWidth < 2 || cellHeight < 2) throw new InvalidOperationException("目标单元格太小或已隐藏，请先调整行高、列宽。");
        double padding = Math.Min(2.0, Math.Min(cellWidth, cellHeight) / 10.0);
        double scale = Math.Min((cellWidth - 2 * padding) / imageWidth, (cellHeight - 2 * padding) / imageHeight);
        double width = imageWidth * scale, height = imageHeight * scale;
        double left = Convert.ToDouble(range.Left) + (cellWidth - width) / 2;
        double top = Convert.ToDouble(range.Top) + (cellHeight - height) / 2;
        object shapes = null, inserted = null;
        string stagingFile = null;
        string stage = "准备插图文件";
        try
        {
            // WPS AddPicture can fail with E_FAIL when given a LocalAppData path.
            // Stage our own file under Documents; embed it and remove the temporary copy.
            string stagingDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "ClipSlim", "InsertCache");
            Directory.CreateDirectory(stagingDirectory);
            stagingFile = Path.Combine(stagingDirectory, "insert-" + Guid.NewGuid().ToString("N") + Path.GetExtension(file));
            File.Copy(file, stagingFile, false);
            stage = "读取工作表图形集合";
            shapes = sheet.Shapes;
            // WPS desktop COM uses points for shape and cell geometry.
            stage = "插入图片";
            inserted = ((dynamic)shapes).AddPicture(stagingFile, 0, -1, left, top, width, height);
            stage = "设置图片大小和位置";
            dynamic picture = inserted;
            picture.Name = "ClipSlim_" + Guid.NewGuid().ToString("N");
            picture.LockAspectRatio = -1;
            picture.Placement = 1; // xlMoveAndSize
            string anchorName = "_ClipSlim_" + Guid.NewGuid().ToString("N");
            object names = null, anchor = null;
            try
            {
                names = sheet.Names;
                string refersTo = "='" + Convert.ToString(sheet.Name).Replace("'", "''") + "'!" + Convert.ToString(range.Address);
                anchor = ((dynamic)names).Add(anchorName, refersTo);
                picture.AlternativeText = "ClipSlimRatio=" + ((double)imageWidth / imageHeight).ToString("R", CultureInfo.InvariantCulture) + ";Anchor=" + anchorName;
            }
            catch { if (anchor != null) { try { ((dynamic)anchor).Delete(); } catch { } } throw; }
            finally { Release(anchor); Release(names); }
            return inserted;
        }
        catch (Exception ex)
        {
            if (inserted != null) { try { ((dynamic)inserted).Delete(); } catch { } Release(inserted); }
            throw new InvalidOperationException(stage + "失败：" + ex.Message, ex);
        }
        finally
        {
            Release(shapes);
            if (stagingFile != null) { try { File.Delete(stagingFile); } catch { } }
        }
    }
    public static int CenterPictures(object sheetObject)
    {
        dynamic sheet = sheetObject;
        if (Convert.ToBoolean(sheet.ProtectDrawingObjects)) return 0;
        object shapes = null;
        int changed = 0;
        try
        {
            shapes = sheet.Shapes;
            int count = Convert.ToInt32(((dynamic)shapes).Count);
            for (int index = 1; index <= count; index++)
            {
                object shape = null, cell = null, range = null;
                try
                {
                    shape = ((dynamic)shapes).Item(index); dynamic p = shape;
                    if (!Convert.ToString(p.Name).StartsWith("ClipSlim_", StringComparison.Ordinal)) continue;
                    string metadata = Convert.ToString(p.AlternativeText) ?? "";
                    // The destination is recorded at insertion. Geometry is never used to
                    // infer it again: WPS may report the previous row while resizing.
                    string[] fields = metadata.Split(';');
                    double ratio;
                    string anchorName = null;
                    foreach (string field in fields) if (field.StartsWith("Anchor=", StringComparison.Ordinal)) anchorName = field.Substring(7);
                    if (fields.Length == 0 || !fields[0].StartsWith("ClipSlimRatio=", StringComparison.Ordinal) || !double.TryParse(fields[0].Substring(14), NumberStyles.Float, CultureInfo.InvariantCulture, out ratio) || ratio <= 0 || anchorName == null || !anchorName.StartsWith("_ClipSlim_", StringComparison.Ordinal)) continue;
                    try
                    {
                        object sheetNames = null, anchor = null;
                        try { sheetNames = sheet.Names; anchor = ((dynamic)sheetNames).Item(anchorName); cell = ((dynamic)anchor).RefersToRange; }
                        finally { Release(anchor); Release(sheetNames); }
                    }
                    catch (COMException) { continue; } // deleted target: do not move to another cell
                    if (cell == null) continue;
                    range = cell; // saved name already refers to the full insertion area
                    dynamic r = range;
                    double cw = Convert.ToDouble(r.Width), ch = Convert.ToDouble(r.Height);
                    if (cw < 2 || ch < 2) continue;
                    double padding = Math.Min(2.0, Math.Min(cw, ch) / 10.0);
                    double w = Math.Min(cw - 2 * padding, (ch - 2 * padding) * ratio), h = w / ratio;
                    double left = Convert.ToDouble(r.Left) + (cw - w) / 2, top = Convert.ToDouble(r.Top) + (ch - h) / 2;
                    if (Math.Abs(Convert.ToDouble(p.Width) - w) < 0.15 && Math.Abs(Convert.ToDouble(p.Height) - h) < 0.15 && Math.Abs(Convert.ToDouble(p.Left) - left) < 0.15 && Math.Abs(Convert.ToDouble(p.Top) - top) < 0.15) continue;
                    p.Placement = 2;
                    p.LockAspectRatio = 0;
                    p.Left = left; p.Top = top;
                    double requestWidth = w, requestHeight = h;
                    for (int attempt = 0; attempt < 4; attempt++)
                    {
                        p.Width = requestWidth; p.Height = requestHeight;
                        double dw = w - Convert.ToDouble(p.Width), dh = h - Convert.ToDouble(p.Height);
                        if (Math.Abs(dw) < 0.1 && Math.Abs(dh) < 0.1) break;
                        requestWidth += dw; requestHeight += dh;
                    }
                    p.LockAspectRatio = -1;
                    // App maintains size; WPS only moves the anchor with cells.
                    if (Trace != null) Trace("CENTER desired=" + w + "x" + h + " actual=" + p.Width + "x" + p.Height + " cell=" + r.Address);
                    changed++;
                }
                finally { if (!Object.ReferenceEquals(cell, range)) Release(range); Release(cell); Release(shape); }
            }
            return changed;
        }
        finally { Release(shapes); }
    }
    public static int CenterActivePictures()
    {
        object app = null, sheet = null;
        try { app = Connect(); sheet = ((dynamic)app).ActiveSheet; return sheet == null ? 0 : CenterPictures(sheet); }
        finally { Release(sheet); Release(app); }
    }
    public static int Test(string directory)
    {
        Directory.CreateDirectory(directory);
        StringBuilder log = new StringBuilder();
        object appObject = null, previous = null, books = null, book = null, sheet = null, single = null, merged = null;
        bool ownsEmptyInstance = false;
        try
        {
            appObject = Activator.CreateInstance(Type.GetTypeFromProgID("ket.Application")); dynamic app = appObject;
            ownsEmptyInstance = Convert.ToInt32(app.Workbooks.Count) == 0;
            Trace = delegate(string line) { log.AppendLine(line); };
            log.AppendLine("STEP: connected to WPS.");
            previous = app.ActiveWorkbook; books = app.Workbooks;
            log.AppendLine("STEP: create test workbook.");
            string testInput = Path.Combine(directory, "fixture-" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(Path.Combine(directory, "fixture.xlsx"), testInput);
            book = ((dynamic)books).Open(Path.GetFullPath(testInput).Replace('/', '\\'));
            log.AppendLine("STEP: read test worksheet.");
            sheet = ((dynamic)book).Worksheets[1]; dynamic ws = sheet;
            log.AppendLine("STEP: configure ordinary cell.");
            single = ws.Range["B2"]; dynamic a = single;
            a.ColumnWidth = 28; a.RowHeight = 100;
            merged = ws.Range["D4:F7"]; dynamic b = merged; b.Merge();
            log.AppendLine("STEP: create image fixture.");
            string actualCache = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ClipSlim", "Clipboard");
            Directory.CreateDirectory(actualCache);
            string file = Path.Combine(actualCache, "轻图-回归-" + Guid.NewGuid().ToString("N") + ".jpg");
            using (Bitmap bitmap = new Bitmap(800, 450))
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.AliceBlue); g.FillEllipse(Brushes.CornflowerBlue, 40, 35, 280, 280);
                using (Font font = new Font("Microsoft YaHei UI", 26)) g.DrawString("轻图 WPS 适配测试", font, Brushes.Black, 35, 350);
                bitmap.Save(file, ImageFormat.Jpeg);
            }
            foreach (object target in new object[] { single, merged })
            {
                log.AppendLine("STEP: insert image.");
                object shape = Insert(sheet, target, file, 800, 450);
                try
                {
                    dynamic s = shape, r = target;
                    double w = Convert.ToDouble(s.Width), h = Convert.ToDouble(s.Height);
                    if (w > Convert.ToDouble(r.Width) || h > Convert.ToDouble(r.Height) || Math.Abs(w / h - 800.0 / 450) > 0.02) throw new Exception("Image sizing / aspect ratio failed");
                    if (Math.Abs(Convert.ToDouble(s.Left) + w / 2 - (Convert.ToDouble(r.Left) + Convert.ToDouble(r.Width) / 2)) > 1) throw new Exception("Horizontal centering failed");
                    if (Math.Abs(Convert.ToDouble(s.Top) + h / 2 - (Convert.ToDouble(r.Top) + Convert.ToDouble(r.Height) / 2)) > 1) throw new Exception("Vertical centering failed");
                    if (Convert.ToInt32(s.Placement) != 1) throw new Exception("Move-and-size property failed");
                    log.AppendLine("PASS: " + r.Address + "; fit and center, aspect ratio preserved; Placement=1; " + w.ToString("0.0") + "x" + h.ToString("0.0") + " points.");
                }
                finally { Release(shape); }
            }
            a.ColumnWidth = 42; a.RowHeight = 155; b.RowHeight = 35;
            int recentered = CenterPictures(sheet);
            object validationShapes = ws.Shapes;
            try
            {
                int i = 1;
                foreach (object target in new object[] { single, merged })
                {
                    object sh = ((dynamic)validationShapes).Item(i++);
                    try
                    {
                        dynamic p = sh, r = target;
                        double width = Convert.ToDouble(p.Width), height = Convert.ToDouble(p.Height);
                        if (Math.Abs(width / height - 800.0 / 450) > 0.02) throw new Exception("Resize changed aspect ratio: " + width + "x" + height + "; metadata=" + p.AlternativeText + "; recentered=" + recentered);
                        if (Math.Abs(Convert.ToDouble(p.Left) + width / 2 - (Convert.ToDouble(r.Left) + Convert.ToDouble(r.Width) / 2)) > 1 || Math.Abs(Convert.ToDouble(p.Top) + height / 2 - (Convert.ToDouble(r.Top) + Convert.ToDouble(r.Height) / 2)) > 1) throw new Exception("Auto-center after row/column resize failed");
                    }
                    finally { Release(sh); }
                }
                if (CenterPictures(sheet) != 0) throw new Exception("Unchanged geometry causes repeated writes");
                log.AppendLine("PASS: row height / column width changed; " + recentered + " images recentered with original aspect ratio; unchanged geometry causes no repeated writes.");
            }
            finally { Release(validationShapes); }
            string workbook = Path.Combine(directory, "wps-fit-test.xlsx");
            ((dynamic)book).SaveAs(workbook, 51);
            log.AppendLine("PASS: saved test workbook with two embedded images: " + workbook);
            File.WriteAllText(Path.Combine(directory, "wps-test.txt"), log.ToString(), Encoding.UTF8);
            return 0;
        }
        catch (Exception ex) { log.AppendLine("FAIL: " + ex); File.WriteAllText(Path.Combine(directory, "wps-test.txt"), log.ToString(), Encoding.UTF8); return 1; }
        finally
        {
            if (book != null) { try { ((dynamic)book).Close(false); } catch { } }
            if (previous != null) { try { ((dynamic)previous).Activate(); } catch { } }
            if (ownsEmptyInstance) { try { ((dynamic)appObject).Quit(); } catch { } }
            Trace = null;
            Release(merged); Release(single); Release(sheet); Release(book); Release(books); Release(previous); Release(appObject);
        }
    }
}

internal static class Compressor
{
    public static byte[] EncodePng(Image image)
    {
        using (MemoryStream stream = new MemoryStream()) { image.Save(stream, ImageFormat.Png); return stream.ToArray(); }
    }

    public static Result Compress(Image input, int quality, int maxEdge, bool png)
    {
        double scale = Math.Min(1.0, (double)maxEdge / Math.Max(input.Width, input.Height));
        int w = Math.Max(1, (int)Math.Round(input.Width * scale));
        int h = Math.Max(1, (int)Math.Round(input.Height * scale));
        using (Bitmap bitmap = new Bitmap(w, h, png ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(png ? Color.Transparent : Color.White);
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(input, new Rectangle(0, 0, w, h), 0, 0, input.Width, input.Height, GraphicsUnit.Pixel);
            }
            byte[] bytes;
            if (png) bytes = EncodePng(bitmap);
            else
            {
                ImageCodecInfo encoder = Array.Find(ImageCodecInfo.GetImageEncoders(), delegate(ImageCodecInfo c) { return c.MimeType == "image/jpeg"; });
                using (EncoderParameters parameters = new EncoderParameters(1))
                using (MemoryStream stream = new MemoryStream())
                {
                    parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);
                    bitmap.Save(stream, encoder, parameters);
                    bytes = stream.ToArray();
                }
            }
            using (MemoryStream stream = new MemoryStream(bytes))
            using (Image decoded = Image.FromStream(stream))
                return new Result { Bytes = bytes, Preview = new Bitmap(decoded), Extension = png ? ".png" : ".jpg" };
        }
    }
}

internal sealed class MainWindow : Form
{
    private readonly PictureBox picture = new PictureBox();
    private readonly Label sourceLabel = new Label();
    private readonly Label resultLabel = new Label();
    private readonly Label status = new Label();
    private readonly NumericUpDown quality = new NumericUpDown();
    private readonly ComboBox edge = new ComboBox();
    private readonly ComboBox format = new ComboBox();
    private readonly ComboBox destination = new ComboBox();
    private readonly OnlineBrowserBridge onlineBrowser = new OnlineBrowserBridge();
    private readonly CheckBox originalView = new CheckBox();
    private readonly CheckBox pin = new CheckBox();
    private readonly CheckBox autoCenter = new CheckBox();
    private readonly System.Windows.Forms.Timer centerTimer = new System.Windows.Forms.Timer();
    private readonly Button compress = new Button();
    private readonly Button copyImage = new Button();
    private readonly Button copyFile = new Button();
    private readonly Button save = new Button();
    private readonly Button insertWps = new Button();
    private readonly Label wpsTarget = new Label();
    private Bitmap source;
    private byte[] sourceBytes;
    private bool estimated;
    private Result result;
    private string cachedFile;
    private bool initialized;

    public MainWindow()
    {
        Text = "轻图 1.7.1 · 在线自动插入";
        Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
        Font = new Font("Microsoft YaHei UI", 9F);
        ClientSize = new Size(560, 500);
        MinimumSize = new Size(576, 489);
        AutoScaleMode = AutoScaleMode.Font;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = Color.FromArgb(245, 247, 250);
        KeyPreview = true;
        AllowDrop = true;

        TableLayoutPanel root = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(12), ColumnCount = 1, RowCount = 10 };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        foreach (int height in new int[] { 29, 37, 29, 40, 0, 26, 31, 40, 28, 43 })
            root.RowStyles.Add(height == 0 ? new RowStyle(SizeType.Percent, 100) : new RowStyle(SizeType.Absolute, height));
        Controls.Add(root);

        Panel header = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty };
        Label title = new Label { Text = "轻图 · 图片压缩", Dock = DockStyle.Fill, Font = new Font(Font.FontFamily, 12F, FontStyle.Bold), ForeColor = Color.FromArgb(28, 45, 67) };
        pin.Text = "置顶";
        pin.Dock = DockStyle.Right;
        pin.Width = 65;
        pin.CheckedChanged += delegate { TopMost = pin.Checked; };
        pin.Checked = true;
        autoCenter.Text = "WPS 自动居中";
        autoCenter.Width = 130;
        autoCenter.Dock = DockStyle.Right;
        autoCenter.Checked = true;
        header.Controls.Add(title);
        header.Controls.Add(autoCenter);
        header.Controls.Add(pin);
        root.Controls.Add(header, 0, 0);
        FlowLayoutPanel input = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        Button paste = Button("读取剪贴板  Ctrl+V", 165);
        paste.Click += delegate { ReadClipboard(); };
        Button open = Button("打开图片…", 100);
        open.Click += delegate { OpenImage(); };
        originalView.Text = "查看原图";
        originalView.AutoSize = true;
        originalView.Margin = new Padding(10, 7, 0, 0);
        originalView.CheckedChanged += delegate { RefreshPreview(); };
        input.Controls.AddRange(new Control[] { paste, open, originalView });
        Button locate = Button("查看 WPS 目标", 110);
        locate.Click += delegate { if (OnlineMode) ShowOnlineHelp(); else LocateWps(); };
        input.Controls.Add(locate);
        root.Controls.Add(input, 0, 1);
        sourceLabel.Text = "复制或截图后读取；也可拖入图片文件。";
        sourceLabel.Dock = DockStyle.Fill;
        sourceLabel.AutoEllipsis = true;
        sourceLabel.ForeColor = Color.FromArgb(81, 96, 112);
        root.Controls.Add(sourceLabel, 0, 2);

        FlowLayoutPanel settings = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        settings.Controls.Add(Caption("格式"));
        format.DropDownStyle = ComboBoxStyle.DropDownList;
        format.Items.AddRange(new object[] { "JPEG", "PNG" });
        format.SelectedIndex = 0;
        format.Width = 78;
        settings.Controls.Add(format);
        settings.Controls.Add(Caption("质量"));
        quality.Minimum = 10; quality.Maximum = 100; quality.Value = 75; quality.Width = 54;
        settings.Controls.Add(quality);
        settings.Controls.Add(Caption("最长边"));
        edge.DropDownStyle = ComboBoxStyle.DropDownList;
        edge.Items.AddRange(new object[] { "原尺寸", "2560", "1920", "1600", "1280", "960", "640" });
        edge.SelectedIndex = 0; edge.Width = 78;
        settings.Controls.Add(edge);
        compress.Text = "压缩"; compress.Size = new Size(65, 28); compress.Margin = new Padding(7, 0, 0, 0);
        Style(compress, true);
        compress.Click += delegate { DoCompress(); };
        settings.Controls.Add(compress);
        root.Controls.Add(settings, 0, 3);

        Panel previewPanel = new Panel { Dock = DockStyle.Fill, BackColor = Color.FromArgb(226, 231, 238), Padding = new Padding(5) };
        picture.Dock = DockStyle.Fill;
        picture.SizeMode = PictureBoxSizeMode.Zoom;
        picture.BackColor = Color.White;
        previewPanel.Controls.Add(picture);
        root.Controls.Add(previewPanel, 0, 4);
        Label previewNote = new Label { Text = "预览压缩结果；PNG 保留透明，JPEG 透明处填白。", Dock = DockStyle.Fill, ForeColor = Color.FromArgb(93, 106, 121), TextAlign = ContentAlignment.MiddleLeft };
        root.Controls.Add(previewNote, 0, 5);
        resultLabel.Dock = DockStyle.Fill;
        resultLabel.Font = new Font(Font.FontFamily, 9F, FontStyle.Bold);
        resultLabel.AutoEllipsis = true;
        resultLabel.TextAlign = ContentAlignment.MiddleLeft;
        root.Controls.Add(resultLabel, 0, 6);
        FlowLayoutPanel actions = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
        insertWps.Text = "插入 WPS 单元格"; insertWps.Size = new Size(155, 30); Style(insertWps, true);
        insertWps.Click += delegate { InsertIntoWps(); };
        copyFile.Text = "复制文件"; copyFile.Size = new Size(95, 30); Style(copyFile, false);
        copyImage.Text = "复制图片"; copyImage.Size = new Size(105, 30); Style(copyImage, false);
        save.Text = "另存为…"; save.Size = new Size(85, 30); Style(save, false);
        copyFile.Click += delegate { Copy(false); };
        copyImage.Click += delegate { Copy(true); };
        save.Click += delegate { SaveResult(); };
        actions.Controls.AddRange(new Control[] { insertWps, copyImage, copyFile, save });
        root.Controls.Add(actions, 0, 7);
        wpsTarget.Dock = DockStyle.Fill;
        wpsTarget.AutoEllipsis = true;
        wpsTarget.Text = "先在 WPS 选中目标单元格，再点击插入；合并单元格也可用。";
        wpsTarget.ForeColor = Color.FromArgb(35, 59, 90);
        Panel targetPanel = new Panel { Dock = DockStyle.Fill };
        destination.DropDownStyle = ComboBoxStyle.DropDownList;
        destination.Items.AddRange(new object[] { "桌面 WPS", "金山在线表格" });
        destination.Width = 145;
        destination.Dock = DockStyle.Left;
        wpsTarget.Padding = new Padding(7, 3, 0, 0);
        targetPanel.Controls.Add(wpsTarget);
        targetPanel.Controls.Add(destination);
        root.Controls.Add(targetPanel, 0, 8);
        destination.SelectedIndexChanged += delegate {
            insertWps.Text = OnlineMode ? "自动插入在线表格" : "插入 WPS 单元格";
            locate.Text = OnlineMode ? "在线设置说明" : "查看 WPS 目标";
            autoCenter.Enabled = !OnlineMode;
            onlineBrowser.Enabled = OnlineMode;
            wpsTarget.Text = OnlineMode ? "先在 Edge 网页点选目标格，再回来点击自动插入。" : "先在 WPS 选中目标单元格。";
            Notify(OnlineMode ? "网页首次设置：插入 → 图片 → 从外部粘贴图片时 → 默认为单元格图片。" : "先在桌面 WPS 选中目标单元格，再点击插入。");
            try { Directory.CreateDirectory(Path.GetDirectoryName(DestinationFile)); File.WriteAllText(DestinationFile, OnlineMode ? "online" : "desktop"); } catch { }
        };
        bool online = false;
        try { online = File.Exists(DestinationFile) && File.ReadAllText(DestinationFile).Trim() == "online"; } catch { }
        destination.SelectedIndex = online ? 1 : 0;
        status.Dock = DockStyle.Fill;
        status.ForeColor = Color.FromArgb(81, 96, 112);
        status.Text = OnlineMode ? "网页首次设置：插入 → 图片 → 从外部粘贴图片时 → 默认为单元格图片。" : "WPS 自动适配：先选中目标单元格，再点“插入 WPS 单元格”。";
        root.Controls.Add(status, 0, 9);

        format.SelectedIndexChanged += delegate { quality.Enabled = format.SelectedIndex == 0; SettingsChanged(); };
        quality.ValueChanged += delegate { SettingsChanged(); };
        edge.SelectedIndexChanged += delegate { SettingsChanged(); };
        KeyDown += delegate(object sender, KeyEventArgs e) { if (e.Control && e.KeyCode == Keys.V) { e.SuppressKeyPress = true; ReadClipboard(); } };
        DragEnter += delegate(object sender, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
        DragDrop += delegate(object sender, DragEventArgs e) { string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; if (files != null && files.Length == 1) LoadFile(files[0]); else Notify("请一次拖入一张图片。"); };
        centerTimer.Interval = 1200;
        centerTimer.Tick += delegate
        {
            if (OnlineMode) { wpsTarget.Text = onlineBrowser.Description; return; }
            if (!autoCenter.Checked || UseWaitCursor) return;
            try
            {
                int count = WpsBridge.CenterActivePictures();
                if (count > 0) Notify("已重新居中 " + count + " 张轻图图片，并按当前单元格保持比例适配。");
            }
            catch { /* WPS may be editing or displaying a dialog; retry on next tick. */ }
        };
        Shown += delegate { centerTimer.Start(); onlineBrowser.Start(); };
        FormClosed += delegate { onlineBrowser.Dispose(); centerTimer.Stop(); centerTimer.Dispose(); picture.Image = null; if (source != null) source.Dispose(); if (result != null) result.Dispose(); };
        initialized = true;
        SetResultEnabled(false);
        compress.Enabled = false;
        StartCollageBridge();
    }

    private void StartCollageBridge()
    {
        Thread worker = new Thread(delegate()
        {
            string sid = WindowsIdentity.GetCurrent().User.Value;
            var security = new PipeSecurity();
            security.AddAccessRule(new PipeAccessRule(WindowsIdentity.GetCurrent().User, PipeAccessRights.FullControl, AccessControlType.Allow));
            while (!IsDisposed)
            {
                try
                {
                    using (var pipe = new NamedPipeServerStream("ClipSlim.Collage.v171." + sid, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.None, 4096, 4096, security))
                    {
                        pipe.WaitForConnection();
                        var reader = new StreamReader(pipe, Encoding.UTF8, false, 4096, true);
                        var writer = new StreamWriter(pipe, new UTF8Encoding(false), 4096, true) { AutoFlush = true };
                        string file = reader.ReadLine();
                        string reply = "ERR|轻图窗口已关闭";
                        Invoke(new Action(delegate()
                        {
                            try
                            {
                                byte[] bytes = File.ReadAllBytes(file);
                                using (var stream = new MemoryStream(bytes))
                                using (var image = Image.FromStream(stream)) SetSource(new Bitmap(image), bytes, false);
                                if (result == null || !insertWps.Enabled) throw new InvalidOperationException(status.Text);
                                if (OnlineMode)
                                {
                                    string page = SendOnlineImage();
                                    reply = "SENT|已向网页发送插入请求：" + page + "；请在网页确认图片。";
                                    return;
                                }
                                using (var target = WpsBridge.ReadTarget())
                                {
                                    object shape = WpsBridge.Insert(target.Sheet, target.Range, EnsureFile(), result.Preview.Width, result.Preview.Height);
                                    WpsBridge.Release(shape);
                                    wpsTarget.Text = "已插入：" + target.Label;
                                    Notify("拼图已压缩并插入 WPS，自动居中已启用。");
                                    reply = "OK|" + target.Label + "（" + SizeText(result.Bytes.Length) + "）";
                                }
                            }
                            catch (Exception ex) { reply = "ERR|" + ex.Message.Replace('\n', ' ').Replace('\r', ' '); Notify(reply.Substring(4)); }
                        }));
                        writer.WriteLine(reply);
                    }
                }
                catch { if (IsDisposed) return; Thread.Sleep(500); }
            }
        });
        worker.IsBackground = true;
        Shown += delegate { worker.Start(); };
    }

    private Label Caption(string text) { return new Label { Text = text, AutoSize = true, Margin = new Padding(4, 5, 3, 0) }; }
    private Button Button(string text, int width) { Button b = new Button { Text = text, Size = new Size(width, 29), Margin = new Padding(0, 0, 8, 0) }; Style(b, false); return b; }
    private void Style(Button b, bool primary) { b.FlatStyle = FlatStyle.Flat; b.FlatAppearance.BorderSize = primary ? 0 : 1; b.BackColor = primary ? Color.FromArgb(35, 99, 224) : Color.White; b.ForeColor = primary ? Color.White : Color.FromArgb(35, 59, 90); b.Cursor = Cursors.Hand; }
    private void Notify(string text) { status.Text = text; }
    private static string SizeText(long length) { return length < 1024 ? length + " B" : length < 1048576 ? (length / 1024.0).ToString("0.0") + " KB" : (length / 1048576.0).ToString("0.00") + " MB"; }
    private void SetResultEnabled(bool enabled) { insertWps.Enabled = copyFile.Enabled = copyImage.Enabled = save.Enabled = enabled; }
    private void LocateWps()
    {
        try { using (WpsTarget t = WpsBridge.ReadTarget()) { wpsTarget.Text = "当前 WPS：" + t.Label; Notify("单元格区域：" + t.Width.ToString("0.0") + " × " + t.Height.ToString("0.0") + " 磅；插入时重新读取当前选中单元格。"); } }
        catch (Exception ex) { wpsTarget.Text = "WPS 未就绪"; Notify("连接失败：" + ex.Message); }
    }
    private void InsertIntoWps()
    {
        if (result == null || !insertWps.Enabled) return;
        if (OnlineMode)
        {
            try { SendOnlineImage(); }
            catch (Exception ex) { Notify("未自动发送：" + ex.Message + " 图片已复制时可手动 Ctrl+V。"); }
            return;
        }
        UseWaitCursor = true;
        try
        {
            using (WpsTarget t = WpsBridge.ReadTarget())
            {
                object shape = WpsBridge.Insert(t.Sheet, t.Range, EnsureFile(), result.Preview.Width, result.Preview.Height);
                WpsBridge.Release(shape);
                wpsTarget.Text = "已插入：" + t.Label;
                Notify("已按比例居中放入单元格，并设为随单元格移动和调整大小；请在 WPS 保存工作簿。");
            }
        }
        catch (Exception ex) { Notify("插入失败：" + ex.Message + "（请退出单元格编辑或对话框后重试）"); }
        finally { UseWaitCursor = false; }
    }
    private bool OnlineMode { get { return destination.SelectedIndex == 1; } }
    private string SendOnlineImage()
    {
        if (!PrepareOnlineImage()) throw new InvalidOperationException(status.Text);
        string page = onlineBrowser.Send();
        wpsTarget.Text = "已发送至：" + page;
        Notify("已自动切回 Edge 并发送粘贴请求，请在网页确认图片。不要连续点击以免重复插入。");
        return page;
    }
    private static string DestinationFile { get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ClipSlim", "destination.txt"); } }
    private void ShowOnlineHelp()
    {
        MessageBox.Show(this,
            "首次在实际使用的 Edge 浏览器中设置：\n插入 → 图片 → 从外部粘贴图片时 → 默认为单元格图片。\n\n在 Edge 金山表格里单击目标格，停留一秒；再点轻图“自动插入在线表格”，或拼图“压缩并发送到表格”。程序会核对目标网址，切回该网页并自动粘贴。请先退出单元格文字编辑状态。\n\n尚未识别目标或网页已切换时，会停止自动操作并提示原因。需要手动粘贴可使用“复制图片”。已发送表示键盘请求已发出，上传是否成功请在网页确认。\n\n网页兼容粘贴使用 PNG 图片，最终体积由金山文档决定。如需精确 JPEG 文件体积，可另存为后通过网页图片入口上传。\n\n只支持桌面版 Edge 中的 HTTPS 金山文档链接；不需要扩展，不读取历史记录。",
            "金山在线表格使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
    private bool PrepareOnlineImage()
    {
        try
        {
            if (!Copy(true)) return false;
            wpsTarget.Text = "待粘贴：金山在线表格（尚未上传）";
            Notify("图片已复制。切回金山网页，选中目标格按 Ctrl+V；首次请设为默认单元格图片。");
            return true;
        }
        catch (Exception ex) { Notify("在线图片准备失败：" + ex.Message); return false; }
    }
    private void SettingsChanged()
    {
        if (!initialized || source == null) return;
        SetResultEnabled(false);
        Notify("参数已更改，请点击“压缩”生成新结果。");
    }
    private void RefreshPreview() { picture.Image = originalView.Checked || result == null ? source : result.Preview; }
    private void OpenImage()
    {
        using (OpenFileDialog dialog = new OpenFileDialog { Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|所有文件|*.*" })
            if (dialog.ShowDialog(this) == DialogResult.OK) LoadFile(dialog.FileName);
    }
    private void LoadFile(string path)
    {
        try
        {
            FileInfo info = new FileInfo(path);
            if (info.Length > 150L * 1024 * 1024) throw new InvalidOperationException("图片文件超过 150 MB，请先缩小图片。");
            byte[] data = File.ReadAllBytes(path);
            using (MemoryStream stream = new MemoryStream(data))
            using (Image image = Image.FromStream(stream))
                SetSource(new Bitmap(image), data, false);
        }
        catch (Exception ex) { Notify("无法打开图片：" + ex.Message); }
    }
    private void ReadClipboard()
    {
        try
        {
            IDataObject data = Clipboard.GetDataObject();
            if (data == null) { Notify("剪贴板为空。"); return; }
            foreach (string name in new string[] { "PNG", "image/png", "image/jpeg", "JFIF" })
            {
                if (!data.GetDataPresent(name, false)) continue;
                object raw = data.GetData(name, false);
                byte[] bytes = raw as byte[];
                MemoryStream memory = raw as MemoryStream;
                if (memory != null) bytes = memory.ToArray();
                if (bytes == null) continue;
                try
                {
                    using (MemoryStream stream = new MemoryStream(bytes))
                    using (Image img = Image.FromStream(stream)) SetSource(new Bitmap(img), bytes, false);
                    return;
                }
                catch (ArgumentException) { }
            }
            if (Clipboard.ContainsImage())
            {
                using (Image image = Clipboard.GetImage())
                {
                    if (image == null) { Notify("无法读取剪贴板图片，请重新复制。"); return; }
                    SetSource(new Bitmap(image), Compressor.EncodePng(image), true);
                    return;
                }
            }
            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = data.GetData(DataFormats.FileDrop) as string[];
                if (files != null && files.Length == 1) { LoadFile(files[0]); return; }
            }
            Notify("剪贴板中没有可读取的单张图片。请先截图或复制一张图片文件。");
        }
        catch (ExternalException) { Notify("剪贴板正被其他程序占用，请稍后重试。"); }
        catch (Exception ex) { Notify("读取失败：" + ex.Message); }
    }
    private void SetSource(Bitmap image, byte[] bytes, bool approximate)
    {
        if ((long)image.Width * image.Height > 80000000) { image.Dispose(); throw new InvalidOperationException("图片超过 8000 万像素，请先缩小图片。"); }
        picture.Image = null;
        if (source != null) source.Dispose();
        if (result != null) { result.Dispose(); result = null; }
        source = image; sourceBytes = bytes; estimated = approximate; cachedFile = null;
        sourceLabel.Text = "原图：" + image.Width + " × " + image.Height + "  ·  " + SizeText(bytes.LongLength) + (estimated ? "（按 PNG 编码估算）" : "");
        compress.Enabled = true;
        DoCompress();
    }
    private void DoCompress()
    {
        if (source == null) return;
        UseWaitCursor = true;
        SetResultEnabled(false);
        try
        {
            int maxEdge = edge.SelectedIndex == 0 ? Math.Max(source.Width, source.Height) : int.Parse(edge.Text);
            Result next = Compressor.Compress(source, (int)quality.Value, maxEdge, format.SelectedIndex == 1);
            picture.Image = null;
            if (result != null) result.Dispose();
            result = next; cachedFile = null;
            double reduction = 1.0 - (double)result.Bytes.Length / sourceBytes.Length;
            resultLabel.Text = "结果：" + result.Preview.Width + " × " + result.Preview.Height + "  ·  " + SizeText(result.Bytes.Length) + "  ·  " + (reduction >= 0 ? "减少 " + reduction.ToString("P1") : "增大 " + (-reduction).ToString("P1")) + (estimated ? "（相对 PNG 估算）" : "");
            RefreshPreview();
            SetResultEnabled(true);
            Notify(reduction < 0 ? "结果比原图大。可降低质量或最长边，也可以选择 PNG 比较。" : "压缩完成。在 WPS 选中目标格，再点“插入 WPS 单元格”，图片会按比例居中适配。");
        }
        catch (Exception ex) { Notify("压缩失败：" + ex.Message); }
        finally { UseWaitCursor = false; }
    }
    private string EnsureFile()
    {
        if (cachedFile != null && File.Exists(cachedFile)) return cachedFile;
        string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ClipSlim", "Clipboard");
        Directory.CreateDirectory(dir);
        cachedFile = Path.Combine(dir, "轻图-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + "-" + Guid.NewGuid().ToString("N").Substring(0, 8) + result.Extension);
        File.WriteAllBytes(cachedFile, result.Bytes);
        return cachedFile;
    }
    private bool Copy(bool asImage)
    {
        if (result == null) return false;
        try
        {
            if (asImage)
            {
                byte[] png = Compressor.EncodePng(result.Preview);
                // CF_DIB is widely accepted by spreadsheets. Publish image formats only:
                // a file-drop or text fallback may make WPS paste a file URI into a cell.
                using (Bitmap opaque = new Bitmap(result.Preview.Width, result.Preview.Height, PixelFormat.Format24bppRgb))
                using (MemoryStream bmp = new MemoryStream())
                using (MemoryStream stream = new MemoryStream(png))
                {
                    using (Graphics g = Graphics.FromImage(opaque))
                    {
                        g.Clear(Color.White);
                        g.DrawImageUnscaled(result.Preview, 0, 0);
                    }
                    opaque.Save(bmp, ImageFormat.Bmp);
                    byte[] file = bmp.ToArray();
                    byte[] dib = new byte[file.Length - 14];
                    Buffer.BlockCopy(file, 14, dib, 0, dib.Length);
                    using (MemoryStream dibStream = new MemoryStream(dib))
                    {
                        DataObject data = new DataObject();
                        data.SetData(DataFormats.Bitmap, false, opaque);
                        data.SetData(DataFormats.Dib, false, dibStream);
                        data.SetData("PNG", false, stream);
                        Clipboard.SetDataObject(data, true, 5, 100);
                    }
                }
                Notify("已复制为图片。切换到 WPS / Excel 按 Ctrl+V 即可；文档内图片体积由接收软件决定。");
            }
            else
            {
                System.Collections.Specialized.StringCollection files = new System.Collections.Specialized.StringCollection();
                files.Add(EnsureFile());
                Clipboard.SetFileDropList(files);
                Notify("已复制压缩文件（" + SizeText(result.Bytes.Length) + "）。可粘贴到聊天软件或文件夹；部分图片输入框不接受文件。");
            }
        }
        catch (Exception ex) { Notify("复制失败，请重试：" + ex.Message); return false; }
        return true;
    }
    private void SaveResult()
    {
        if (result == null) return;
        using (SaveFileDialog dialog = new SaveFileDialog { Filter = result.Extension == ".jpg" ? "JPEG 图片|*.jpg" : "PNG 图片|*.png", FileName = "轻图-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + result.Extension })
        {
            if (dialog.ShowDialog(this) != DialogResult.OK) return;
            try { File.WriteAllBytes(dialog.FileName, result.Bytes); Notify("已保存：" + dialog.FileName); }
            catch (Exception ex) { Notify("保存失败：" + ex.Message); }
        }
    }

    internal int IntegrationTest(string directory)
    {
        Directory.CreateDirectory(directory);
        StringBuilder log = new StringBuilder();
        try
        {
            using (Bitmap fixture = new Bitmap(1280, 720))
            {
                using (Graphics g = Graphics.FromImage(fixture))
                {
                    g.Clear(Color.White);
                    g.FillRectangle(Brushes.CornflowerBlue, 30, 30, 600, 400);
                    g.DrawString("轻图剪贴板测试 ABC 123", Font, Brushes.Black, 40, 460);
                }
                Clipboard.SetImage(fixture);
            }
            ReadClipboard();
            if (source == null || result == null || source.Width != 1280) throw new Exception("Clipboard bitmap input failed");
            log.AppendLine("PASS: read Windows bitmap clipboard and compress.");
            edge.SelectedItem = "640";
            if (insertWps.Enabled || copyImage.Enabled || copyFile.Enabled || save.Enabled) throw new Exception("Stale output still enabled after settings change");
            DoCompress();
            if (result.Preview.Width != 640 || result.Preview.Height != 360) throw new Exception("UI resize settings failed");
            log.AppendLine("PASS: changed settings invalidate output; recompression applies longest edge.");
            Copy(false);
            System.Collections.Specialized.StringCollection files = Clipboard.GetFileDropList();
            if (files.Count != 1 || !File.Exists(files[0])) throw new Exception("Clipboard file missing");
            byte[] onDisk = File.ReadAllBytes(files[0]);
            if (Convert.ToBase64String(onDisk) != Convert.ToBase64String(result.Bytes)) throw new Exception("File clipboard payload mismatch");
            ReadClipboard();
            if (source == null || estimated || source.Width != 640) throw new Exception("File clipboard input failed");
            log.AppendLine("PASS: copy compressed file; exact encoded bytes on disk; reread file clipboard.");
            Copy(true);
            using (Image pasted = Clipboard.GetImage())
                if (pasted == null || pasted.Width != 640 || pasted.Height != 360) throw new Exception("Image clipboard output failed");
            if (!Clipboard.GetDataObject().GetDataPresent("PNG", false)) throw new Exception("Clipboard PNG format missing");
            IDataObject clipboard = Clipboard.GetDataObject();
            if (!clipboard.GetDataPresent(DataFormats.Dib, false)) throw new Exception("Spreadsheet DIB image format missing");
            if (clipboard.GetDataPresent(DataFormats.FileDrop) || clipboard.GetDataPresent(DataFormats.UnicodeText)) throw new Exception("Image clipboard contains file or text fallback");
            log.AppendLine("PASS: spreadsheet DIB format present; no file-drop or text path in image clipboard.");
            ReadClipboard();
            if (source == null || estimated || source.Width != 640) throw new Exception("PNG clipboard input failed: " + status.Text + "; estimated=" + estimated + "; source width=" + (source == null ? 0 : source.Width));
            Copy(true);
            log.AppendLine("PASS: copy real image plus PNG; reread PNG clipboard.");
            int priorDestination = destination.SelectedIndex;
            destination.SelectedIndex = 1;
            if (insertWps.Text != "自动插入在线表格" || autoCenter.Enabled) throw new Exception("Online mode controls failed");
            if (!PrepareOnlineImage()) throw new Exception("Online clipboard preparation failed: " + status.Text);
            if (!wpsTarget.Text.Contains("尚未上传")) throw new Exception("Online mode incorrectly claims completed insertion");
            IDataObject onlineClipboard = Clipboard.GetDataObject();
            if (!onlineClipboard.GetDataPresent("PNG", false) || onlineClipboard.GetDataPresent(DataFormats.FileDrop) || onlineClipboard.GetDataPresent(DataFormats.UnicodeText)) throw new Exception("Unsafe web image clipboard formats");
            destination.SelectedIndex = priorDestination;
            log.AppendLine("PASS: online mode uses image-only clipboard; disables desktop centering; reports pending paste honestly.");
            File.WriteAllText(Path.Combine(directory, "integration-test.txt"), log.ToString(), Encoding.UTF8);
            return 0;
        }
        catch (Exception ex) { log.AppendLine("FAIL: " + ex); File.WriteAllText(Path.Combine(directory, "integration-test.txt"), log.ToString(), Encoding.UTF8); return 1; }
    }

    internal int LayoutTest(string report)
    {
        try
        {
            Show();
            Application.DoEvents();
            if (!TopMost || !pin.Checked) throw new Exception("Default pin failed");
            pin.Checked = false;
            if (TopMost) throw new Exception("Unpin failed");
            pin.Checked = true;
            foreach (Size size in new Size[] { new Size(560, 500), new Size(560, 450), new Size(760, 620) })
            {
                ClientSize = size;
                PerformLayout();
                Application.DoEvents();
                foreach (Control control in new Control[] { compress, insertWps, copyImage, copyFile, save, pin, originalView, format, quality, edge })
                {
                    if (!control.Visible || control.Right > control.Parent.ClientSize.Width || control.Bottom > control.Parent.ClientSize.Height)
                        throw new Exception("Control clipped: " + control.Text + " at " + ClientSize);
                }
                if (picture.Height < 60) throw new Exception("Preview too small");
            }
            File.WriteAllText(report, "PASS: compact layout at 560x500, 560x450, 760x620; controls fit; preview visible; topmost on/off toggle.", Encoding.UTF8);
            return 0;
        }
        catch (Exception ex) { File.WriteAllText(report, "FAIL: " + ex, Encoding.UTF8); return 1; }
    }
}

internal static class Tests
{
    private static void Check(bool condition, string message) { if (!condition) throw new Exception(message); }
    public static int Run(string directory)
    {
        Directory.CreateDirectory(directory);
        StringBuilder log = new StringBuilder();
        try
        {
            using (Bitmap input = new Bitmap(1920, 1080, PixelFormat.Format32bppArgb))
            {
                Random random = new Random(17);
                for (int y = 0; y < input.Height; y++)
                    for (int x = 0; x < input.Width; x++)
                        input.SetPixel(x, y, Color.FromArgb(255, (x / 8 + random.Next(12)) % 256, (y / 5 + random.Next(12)) % 256, (x / 12 + y / 8) % 256));
                byte[] original = Compressor.EncodePng(input);
                using (Result small = Compressor.Compress(input, 75, 1280, false))
                {
                    Check(small.Preview.Width == 1280 && small.Preview.Height == 720, "Resize aspect ratio");
                    Check(small.Bytes.Length < original.Length, "JPEG size reduction");
                    Check(small.Bytes[0] == 0xff && small.Bytes[1] == 0xd8, "JPEG signature");
                    File.WriteAllBytes(Path.Combine(directory, "compressed.jpg"), small.Bytes);
                    log.AppendLine("PASS: JPEG, resize 1920x1080 -> 1280x720; " + original.Length + " -> " + small.Bytes.Length + " bytes.");
                }
                using (Result full = Compressor.Compress(input, 90, 2560, false))
                    Check(full.Preview.Width == 1920 && full.Preview.Height == 1080, "No upscaling");
                using (Result low = Compressor.Compress(input, 20, 1920, false))
                using (Result high = Compressor.Compress(input, 95, 1920, false))
                    Check(low.Bytes.Length < high.Bytes.Length, "JPEG quality control");
                log.AppendLine("PASS: no upscaling; JPEG quality affects output size.");
            }
            using (Bitmap transparent = new Bitmap(100, 60, PixelFormat.Format32bppArgb))
            {
                using (Graphics g = Graphics.FromImage(transparent)) { g.Clear(Color.Transparent); g.FillRectangle(Brushes.Red, 30, 20, 30, 20); }
                using (Result png = Compressor.Compress(transparent, 75, 100, true))
                {
                    Check(png.Preview.GetPixel(0, 0).A == 0, "PNG transparency");
                    File.WriteAllBytes(Path.Combine(directory, "transparent.png"), png.Bytes);
                }
                using (Result jpg = Compressor.Compress(transparent, 75, 100, false))
                    Check(jpg.Preview.GetPixel(0, 0).R > 245, "JPEG white background");
                log.AppendLine("PASS: PNG transparency and JPEG white background.");
            }
            using (Bitmap pixel = new Bitmap(1, 1))
            using (Result tiny = Compressor.Compress(pixel, 75, 640, false))
                Check(tiny.Preview.Width == 1 && tiny.Preview.Height == 1, "One-pixel image");
            log.AppendLine("PASS: one-pixel image.");
            File.WriteAllText(Path.Combine(directory, "self-test.txt"), log.ToString(), Encoding.UTF8);
            return 0;
        }
        catch (Exception ex) { log.AppendLine("FAIL: " + ex); File.WriteAllText(Path.Combine(directory, "self-test.txt"), log.ToString(), Encoding.UTF8); return 1; }
    }
}
