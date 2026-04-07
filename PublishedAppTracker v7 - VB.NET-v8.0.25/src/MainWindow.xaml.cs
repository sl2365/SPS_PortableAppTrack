using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;

namespace PublishedAppTracker
{
    public partial class MainWindow : Window
    {
        private string appDir;
        private string categoriesPath;
        private string settingsPath;
        private string webViewPath;
        private bool isHorizontalLayout = false;
        private List<TrackItem> currentItems = new List<TrackItem>();
        private string lastSortColumn = "";
        private bool lastSortAscending = true;

        // Shared controls - created once in code
        private DockPanel categoryPanel;
        private TreeView categoryTree;
        private ListView itemList;
        private DockPanel trackSettingsPanel;
        private TabControl rightTabs;
        private Microsoft.Web.WebView2.Wpf.WebView2 webView;

        // Track settings fields
        private TextBox editName;
        private TextBox editTrackURL;
        private TextBox editStartString;
        private TextBox editStopString;
        private TextBox editDownloadURL;
        private TextBox editVersion;
        private TextBlock editLatestVersion;
        private TextBox editReleaseDate;
        private TextBox editPublisherName;
        private TextBox editSuiteName;
        private TrackItem currentTrackItem = null;
        private string currentCategoryPath = null;
        private TextBlock startPositionText;
        private TextBlock stopPositionText;
        private Button btnUpdateVersion;
        private bool suppressAutoDownload = false;
        private bool blockCookiePopups = true;
        private TextBox editEditorPath;

        // Source tab fields
        private RichTextBox sourceView;
        private TextBox findString;
        private TextBlock findCount;
        private string extensionsPath;

        // Download + search state
        private static readonly HttpClient httpClient = new HttpClient(
            new HttpClientHandler
            {
                AutomaticDecompression = System.Net.DecompressionMethods.GZip |
                                         System.Net.DecompressionMethods.Deflate |
                                         System.Net.DecompressionMethods.Brotli
            });
        private const string UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:123.0) Gecko/20100101 Firefox/123.0";
        private string currentSource = "";
        private List<int> searchHitPositions = new List<int>();
        private int currentSearchIndex = -1;

		// Window settings
        private WindowSettings windowSettings;
        private string windowSettingsPath;
        private StackPanel columnCheckboxPanel;
  
        // Theme
		private ThemeSettings currentTheme;
		private string currentThemePath;
		private ThemeSettings previewTheme;

        // Window constructor
        public MainWindow()
        {
            InitializeComponent();

            appDir = Path.GetDirectoryName(
                System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName);

            categoriesPath = Path.Combine(appDir, "Categories");
            settingsPath = Path.Combine(appDir, "Settings");
            webViewPath = Path.Combine(appDir, "WebView");
            extensionsPath = Path.Combine(appDir, "Extensions");

            if (!Directory.Exists(categoriesPath))
                Directory.CreateDirectory(categoriesPath);
            if (!Directory.Exists(settingsPath))
                Directory.CreateDirectory(settingsPath);
            if (!Directory.Exists(webViewPath))
                Directory.CreateDirectory(webViewPath);
            if (!Directory.Exists(extensionsPath))
                Directory.CreateDirectory(extensionsPath);

            // Load window settings
            windowSettingsPath = Path.Combine(settingsPath, "window.xml");
            windowSettings = WindowSettings.Load(windowSettingsPath);

            // Apply window position and size
            this.Left = windowSettings.WindowLeft;
            this.Top = windowSettings.WindowTop;
            this.Width = windowSettings.WindowWidth;
            this.Height = windowSettings.WindowHeight;
            if (windowSettings.IsMaximized)
                this.WindowState = WindowState.Maximized;

            isHorizontalLayout = windowSettings.IsHorizontalLayout;

            // Apply layout visibility to match saved setting
            if (isHorizontalLayout)
            {
			    layoutVertical.Visibility = Visibility.Collapsed;
			    layoutHorizontal.Visibility = Visibility.Visible;
			    btnToggleLayout.Content = "\xeca5";
			    btnToggleLayout.FontFamily = new FontFamily("Segoe Fluent Icons");
			    btnToggleLayout.FontSize = 16;
			    btnToggleLayout.ToolTip = "Switch to Vertical Layout";
            }

			currentTheme = ThemeSettings.GetDefaultLight();
			currentThemePath = "";
			if (!string.IsNullOrEmpty(windowSettings.ActiveThemePath))
			{
			    string themePath = Path.Combine(settingsPath, windowSettings.ActiveThemePath);
			    if (File.Exists(themePath))
			    {
			        currentTheme = ThemeSettings.Load(themePath);
			        currentThemePath = themePath;
			    }
			}
			previewTheme = currentTheme.Clone();

			BuildSharedControls();
            PlaceControlsInLayout();
            ApplyTheme(currentTheme);
            SetupWebViewAutoZoom();
            // Pre-initialise WebView2 with extensions enabled
            InitWebView();
            RefreshCategoryTree();

            // Select first category on startup
            if (categoryTree.Items.Count > 0)
            {
                TreeViewItem firstCat = categoryTree.Items[0] as TreeViewItem;
                if (firstCat != null)
                    firstCat.IsSelected = true;
            }

            // Apply splitter positions after layout is ready
            this.Loaded += (s, ev) =>
            {
                ApplySplitterPositions();
            };

            // Save settings on close
            this.Closing += (s, ev) =>
            {
                SaveWindowSettings();
            };

            statusFile.Text = "Ready — " + appDir;
        }
		
        // ============================
        // Build Shared Controls Once
        // ============================

        private void SetupListViewHorizontalScrolling()
        {
            // Normal scroll wheel: horizontal when no vertical scrollbar, or with Shift held
            itemList.PreviewMouseWheel += (s, e) =>
            {
                ScrollViewer sv = FindScrollViewer(itemList);
                if (sv == null) return;

                bool hasVerticalScroll = sv.ScrollableHeight > 0;

                if (System.Windows.Input.Keyboard.IsKeyDown(System.Windows.Input.Key.LeftShift) ||
                    System.Windows.Input.Keyboard.IsKeyDown(System.Windows.Input.Key.RightShift))
                {
                    if (e.Delta > 0)
                        sv.LineLeft();
                    else
                        sv.LineRight();

                    e.Handled = true;
                    return;
                }

                if (!hasVerticalScroll && sv.ScrollableWidth > 0)
                {
                    if (e.Delta > 0)
                        sv.LineLeft();
                    else
                        sv.LineRight();

                    e.Handled = true;
                }
            };

            // Tilt wheel: hook Windows message for WM_MOUSEHWHEEL
            itemList.Loaded += (s, e) =>
            {
                var source = PresentationSource.FromVisual(itemList) as System.Windows.Interop.HwndSource;
                if (source != null)
                {
                    source.AddHook(ListViewTiltWheelHook);
                }
            };
        }

        private const int WM_MOUSEHWHEEL = 0x020E;

        private IntPtr ListViewTiltWheelHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_MOUSEHWHEEL)
            {
                // Check if mouse is over the ListView
                Point mousePos = System.Windows.Input.Mouse.GetPosition(itemList);
                if (mousePos.X >= 0 && mousePos.Y >= 0 &&
                    mousePos.X <= itemList.ActualWidth && mousePos.Y <= itemList.ActualHeight)
                {
                    ScrollViewer sv = FindScrollViewer(itemList);
                    if (sv != null && sv.ScrollableWidth > 0)
                    {
                        int delta = (short)((wParam.ToInt64() >> 16) & 0xFFFF);

                        if (delta > 0)
                            sv.LineRight();
                        else
                            sv.LineLeft();

                        handled = true;
                    }
                }
            }

            return IntPtr.Zero;
        }

        private ScrollViewer FindScrollViewer(DependencyObject obj)
        {
            if (obj is ScrollViewer sv)
                return sv;

            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                ScrollViewer result = FindScrollViewer(child);
                if (result != null)
                    return result;
            }

            return null;
        }

        private void BuildSharedControls()
        {
            // --- Category Tree Panel ---
            categoryPanel = new DockPanel();

            StackPanel catButtons = new StackPanel();
            catButtons.Orientation = Orientation.Horizontal;
            catButtons.Margin = new Thickness(4);
            DockPanel.SetDock(catButtons, Dock.Bottom);

            Button btnAdd = new Button();
            btnAdd.Content = "\xecc8";
            btnAdd.FontFamily = new FontFamily("Segoe Fluent Icons");
    		btnAdd.FontSize = 14;
            btnAdd.Width = 30;
            btnAdd.Height = 25;
            btnAdd.ToolTip = "New Category";
            btnAdd.Margin = new Thickness(4, 0, 4, 0);
            btnAdd.Click += AddCategory_Click;
            catButtons.Children.Add(btnAdd);

            Button btnRemove = new Button();
            btnRemove.Content = "\xecc9";
            btnRemove.FontFamily = new FontFamily("Segoe Fluent Icons");
    		btnRemove.FontSize = 14;
            btnRemove.Width = 30;
            btnRemove.Height = 25;
            btnRemove.ToolTip = "Delete Category";
            btnRemove.Click += RemoveCategory_Click;
            catButtons.Children.Add(btnRemove);

            Button btnOpenFolder = new Button();
            btnOpenFolder.Content = "\xe838";
            btnOpenFolder.FontFamily = new FontFamily("Segoe Fluent Icons");
    		btnOpenFolder.FontSize = 14;
            btnOpenFolder.Width = 30;
            btnOpenFolder.Height = 25;
            btnOpenFolder.ToolTip = "Open category folder";
            btnOpenFolder.Margin = new Thickness(4, 0, 0, 0);
            btnOpenFolder.Click += OpenCategoryFolder_Click;
            catButtons.Children.Add(btnOpenFolder);

            categoryPanel.Children.Add(catButtons);

            categoryTree = new TreeView();
            categoryTree.Margin = new Thickness(4);
            categoryTree.SelectedItemChanged += CategoryTree_Selected;
            categoryPanel.Children.Add(categoryTree);

            // --- Item List ---
            itemList = new ListView();
            itemList.Margin = new Thickness(4);
            itemList.SelectionChanged += ItemList_SelectionChanged;
            itemList.AddHandler(GridViewColumnHeader.ClickEvent,
                new RoutedEventHandler(ColumnHeader_Click));

            GridView gridView = new GridView();

            GridViewColumn checkCol = new GridViewColumn();
            checkCol.Width = 30;
            FrameworkElementFactory headerCheckbox = new FrameworkElementFactory(typeof(CheckBox));
            headerCheckbox.AddHandler(CheckBox.ClickEvent, new RoutedEventHandler(SelectAll_Click));
            checkCol.Header = new GridViewColumnHeader();
            checkCol.HeaderTemplate = new DataTemplate();
            checkCol.HeaderTemplate.VisualTree = headerCheckbox;
            FrameworkElementFactory cellCheckbox = new FrameworkElementFactory(typeof(CheckBox));
            cellCheckbox.SetBinding(CheckBox.IsCheckedProperty,
                new System.Windows.Data.Binding("IsSelected"));
            cellCheckbox.SetValue(CheckBox.HorizontalAlignmentProperty,
                HorizontalAlignment.Center);
            checkCol.CellTemplate = new DataTemplate();
            checkCol.CellTemplate.VisualTree = cellCheckbox;
            gridView.Columns.Add(checkCol);

            GridViewColumn statusCol = new GridViewColumn();
            statusCol.Header = "St";
            statusCol.Width = 30;
            statusCol.CellTemplate = CreateStatusCellTemplate();
            gridView.Columns.Add(statusCol);

            // Build columns from saved settings or defaults
            List<ColumnSetting> colSettings = windowSettings.ColumnSettings;
            if (colSettings.Count == 0)
                colSettings = WindowSettings.GetDefaultColumns();

            // Sort by DisplayIndex for correct order
            colSettings.Sort((a, b) => a.DisplayIndex.CompareTo(b.DisplayIndex));

            foreach (ColumnSetting cs in colSettings)
            {
                if (cs.Visible)
                {
                    string statusBinding = GetStatusBindingForColumn(cs.Binding);
                    if (statusBinding != null)
                        gridView.Columns.Add(MakeColumn(cs.Header, cs.Binding, cs.Width, statusBinding));
                    else
                        gridView.Columns.Add(MakeColumn(cs.Header, cs.Binding, cs.Width));
                }
            }

            itemList.View = gridView;
            ContextMenu listContextMenu = new ContextMenu();

            MenuItem ctxNewTrack = new MenuItem();
            ctxNewTrack.Header = "New Track";
            ctxNewTrack.Click += MenuNewTrack_Click;
            listContextMenu.Items.Add(ctxNewTrack);

            listContextMenu.Items.Add(new Separator());

            MenuItem ctxOpenFile = new MenuItem();
            ctxOpenFile.Header = "Open File in Editor";
            ctxOpenFile.Click += OpenFileInEditor_Click;
            listContextMenu.Items.Add(ctxOpenFile);

            listContextMenu.Items.Add(new Separator());

            MenuItem ctxSave = new MenuItem();
            ctxSave.Header = "Save Track";
            ctxSave.Click += SaveTrack_Click;
            listContextMenu.Items.Add(ctxSave);

            MenuItem ctxDelete = new MenuItem();
            ctxDelete.Header = "Delete Track";
            ctxDelete.Click += DeleteTrack_Click;
            listContextMenu.Items.Add(ctxDelete);

            itemList.ContextMenu = listContextMenu;
            SetupListViewHorizontalScrolling();

            // --- Track Settings Panel (Panel 3) ---
            BuildTrackSettingsPanel();

            // --- Tabbed Pane (Panel 4) ---
            BuildTabbedPane();
        }

        private void BuildTrackSettingsPanel()
        {
            trackSettingsPanel = new DockPanel();

            ScrollViewer scrollArea = new ScrollViewer();
            scrollArea.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            DockPanel outerPanel = new DockPanel();

            // Track Settings toolbar
            ToolBarTray trackToolBarTray = new ToolBarTray();
            DockPanel.SetDock(trackToolBarTray, Dock.Top);

            ToolBar trackToolBar = new ToolBar();
            trackToolBar.Band = 0;
            trackToolBar.BandIndex = 0;

            TextBlock title = new TextBlock();
            title.Text = "Track Settings";
            title.FontWeight = FontWeights.Bold;
            title.FontSize = 13;
            title.VerticalAlignment = VerticalAlignment.Center;
            title.Margin = new Thickness(0, 0, 8, 0);
            trackToolBar.Items.Add(title);

            trackToolBar.Items.Add(new Separator());

			Button btnSaveTrack = new Button();
			btnSaveTrack.Content = "\uE74E";
			btnSaveTrack.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnSaveTrack.FontSize = 17;
			btnSaveTrack.Padding = new Thickness(6, 2, 6, 2);
			btnSaveTrack.ToolTip = "Save track settings";
			btnSaveTrack.Click += SaveTrack_Click;
			trackToolBar.Items.Add(btnSaveTrack);
			
            Button btnDeleteTrack = new Button();
            btnDeleteTrack.Content = "\uE74D";
			btnDeleteTrack.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnDeleteTrack.FontSize = 16;
            btnDeleteTrack.Padding = new Thickness(6, 2, 6, 2);
            btnDeleteTrack.Margin = new Thickness(2, 0, 0, 0);
            btnDeleteTrack.ToolTip = "Delete this track";
            btnDeleteTrack.Click += DeleteTrack_Click;
            trackToolBar.Items.Add(btnDeleteTrack);

            trackToolBar.Items.Add(new Separator());

            Button btnGoStart = new Button();
            btnGoStart.Content = "\uE768";
			btnGoStart.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnGoStart.FontSize = 16;
            btnGoStart.Padding = new Thickness(6, 2, 6, 2);
            btnGoStart.ToolTip = "Go to Start String in source";
            btnGoStart.Click += GoToStartString_Click;
            trackToolBar.Items.Add(btnGoStart);

            Button btnGoStop = new Button();
            btnGoStop.Content = "\uE71A";
			btnGoStop.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnGoStop.FontSize = 16;
            btnGoStop.Padding = new Thickness(6, 2, 6, 2);
            btnGoStop.Margin = new Thickness(2, 0, 0, 0);
            btnGoStop.ToolTip = "Go to Stop String in source";
            btnGoStop.Click += GoToStopString_Click;
            trackToolBar.Items.Add(btnGoStop);

            trackToolBar.Items.Add(new Separator());

            Button btnDownloadToolbar = new Button();
            btnDownloadToolbar.Content = "\uE896";
			btnDownloadToolbar.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnDownloadToolbar.FontSize = 16;
            btnDownloadToolbar.Padding = new Thickness(6, 2, 6, 2);
            btnDownloadToolbar.ToolTip = "Download page source";
            btnDownloadToolbar.Click += DownloadPage_Click;
            trackToolBar.Items.Add(btnDownloadToolbar);

            Button btnBrowserToolbar = new Button();
            btnBrowserToolbar.Content = "\uE774";
			btnBrowserToolbar.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnBrowserToolbar.FontSize = 16;
            btnBrowserToolbar.Padding = new Thickness(6, 2, 6, 2);
            btnBrowserToolbar.Margin = new Thickness(2, 0, 0, 0);
            btnBrowserToolbar.ToolTip = "Open in browser";
            btnBrowserToolbar.Click += OpenBrowser_Click;
            trackToolBar.Items.Add(btnBrowserToolbar);

            trackToolBar.Items.Add(new Separator());

            btnUpdateVersion = new Button();
            btnUpdateVersion.Content = "\uE898";
			btnUpdateVersion.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnUpdateVersion.FontSize = 16;
            btnUpdateVersion.Padding = new Thickness(6, 2, 6, 2);
            btnUpdateVersion.ToolTip = "Copy Latest Version to Version field";
            btnUpdateVersion.IsEnabled = false;
            btnUpdateVersion.Click += UpdateVersion_Click;
            trackToolBar.Items.Add(btnUpdateVersion);

            trackToolBarTray.ToolBars.Add(trackToolBar);

            // Fields
            StackPanel fieldsPanel = new StackPanel();
            fieldsPanel.Margin = new Thickness(8);

            editName = AddSettingsField(fieldsPanel, "Program Name:");

            // Track URL with buttons
            TextBlock urlLabel = new TextBlock();
            urlLabel.Text = "Track URL:";
            urlLabel.Margin = new Thickness(0, 0, 0, 4);
            fieldsPanel.Children.Add(urlLabel);

            DockPanel urlDock = new DockPanel();
            urlDock.Margin = new Thickness(0, 0, 0, 8);

            editTrackURL = new TextBox();
            urlDock.Children.Add(editTrackURL);
            fieldsPanel.Children.Add(urlDock);

            // Start String with Go To button
            DockPanel startLabelDock = new DockPanel();
            startLabelDock.Margin = new Thickness(0, 0, 0, 4);

            startPositionText = new TextBlock();
            startPositionText.Text = "";
            startPositionText.Foreground = Brushes.Gray;
            startPositionText.FontSize = 11;
            startPositionText.HorizontalAlignment = HorizontalAlignment.Right;
            DockPanel.SetDock(startPositionText, Dock.Right);
            startLabelDock.Children.Add(startPositionText);

            TextBlock startLabel = new TextBlock();
            startLabel.Text = "Start String:";
            startLabelDock.Children.Add(startLabel);

            fieldsPanel.Children.Add(startLabelDock);

            DockPanel startDock = new DockPanel();
            startDock.Margin = new Thickness(0, 0, 0, 8);

            editStartString = new TextBox();
            editStartString.Height = 60;
            editStartString.AcceptsReturn = true;
            editStartString.TextWrapping = TextWrapping.Wrap;
            editStartString.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            startDock.Children.Add(editStartString);
            fieldsPanel.Children.Add(startDock);

            // Stop String with Go To button
            DockPanel stopLabelDock = new DockPanel();
            stopLabelDock.Margin = new Thickness(0, 0, 0, 4);

            stopPositionText = new TextBlock();
            stopPositionText.Text = "";
            stopPositionText.Foreground = Brushes.Gray;
            stopPositionText.FontSize = 11;
            stopPositionText.HorizontalAlignment = HorizontalAlignment.Right;
            DockPanel.SetDock(stopPositionText, Dock.Right);
            stopLabelDock.Children.Add(stopPositionText);

            TextBlock stopLabel = new TextBlock();
            stopLabel.Text = "Stop String:";
            stopLabelDock.Children.Add(stopLabel);

            fieldsPanel.Children.Add(stopLabelDock);

            DockPanel stopDock = new DockPanel();
            stopDock.Margin = new Thickness(0, 0, 0, 8);

            editStopString = new TextBox();
            editStopString.Height = 60;
            editStopString.AcceptsReturn = true;
            editStopString.TextWrapping = TextWrapping.Wrap;
            editStopString.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            stopDock.Children.Add(editStopString);
            fieldsPanel.Children.Add(stopDock);
            editDownloadURL = AddSettingsField(fieldsPanel, "Download URL:");
            editVersion = AddSettingsField(fieldsPanel, "Version:", 150);

            // Read-only latest version display with update button
            TextBlock latestLabel = new TextBlock();
            latestLabel.Text = "Latest Version (auto-detected):";
            latestLabel.Margin = new Thickness(0, 0, 0, 4);
            fieldsPanel.Children.Add(latestLabel);

            DockPanel latestDock = new DockPanel();
            latestDock.Margin = new Thickness(0, 0, 0, 8);

            TextBlock latestDisplay = new TextBlock();
            latestDisplay.Text = "";
            latestDisplay.FontWeight = FontWeights.Bold;
            latestDisplay.FontSize = 14;
            latestDisplay.VerticalAlignment = VerticalAlignment.Center;
            latestDock.Children.Add(latestDisplay);
            editLatestVersion = latestDisplay;

            fieldsPanel.Children.Add(latestDock);

            // Metadata fields
            editReleaseDate = AddSettingsField(fieldsPanel, "Release Date:");
            editPublisherName = AddSettingsField(fieldsPanel, "Publisher Name:");
            editSuiteName = AddSettingsField(fieldsPanel, "Suite Name:");

            outerPanel.Children.Add(fieldsPanel);
            scrollArea.Content = outerPanel;

            DockPanel.SetDock(trackToolBarTray, Dock.Top);
            trackSettingsPanel.Children.Add(trackToolBarTray);
            trackSettingsPanel.Children.Add(scrollArea);
        }

        private void BuildTabbedPane()
        {
            rightTabs = new TabControl();
            rightTabs.Margin = new Thickness(4);

            // Tab 1: Source
			TabItem sourceTab = new TabItem();
			sourceTab.Header = "📄 Source";
			DockPanel sourcePanel = new DockPanel();

			// Source tab toolbar
			ToolBarTray sourceToolBarTray = new ToolBarTray();
			DockPanel.SetDock(sourceToolBarTray, Dock.Top);

			ToolBar sourceToolBar = new ToolBar();
			sourceToolBar.Band = 0;
			sourceToolBar.BandIndex = 0;

			findString = new TextBox();
			findString.Width = 300;
			findString.Margin = new Thickness(0, 0, 4, 0);
			findString.KeyDown += (s, ev) =>
			{
			    if (ev.Key == System.Windows.Input.Key.Enter)
			    {
			        SearchDown_Click(findString, new RoutedEventArgs());
			    }
			};
			sourceToolBar.Items.Add(findString);

			Button btnSearchDown = new Button();
			StackPanel spDown = new StackPanel();
			spDown.Orientation = Orientation.Horizontal;

			TextBlock iconSearchDown = new TextBlock();
			iconSearchDown.Text = "\uE721";
			iconSearchDown.FontFamily = new FontFamily("Segoe Fluent Icons");
			iconSearchDown.FontSize = 16;
			iconSearchDown.VerticalAlignment = VerticalAlignment.Center;
			spDown.Children.Add(iconSearchDown);

			TextBlock iconArrowDown = new TextBlock();
			iconArrowDown.Text = "\uE760";
			iconArrowDown.FontFamily = new FontFamily("Segoe Fluent Icons");
			iconArrowDown.FontSize = 16;
			iconArrowDown.VerticalAlignment = VerticalAlignment.Center;
			iconArrowDown.RenderTransformOrigin = new Point(0.5, 0.5);
			iconArrowDown.RenderTransform = new RotateTransform(270);
			spDown.Children.Add(iconArrowDown);

			btnSearchDown.Content = spDown;
			btnSearchDown.Width = 48;
			btnSearchDown.Padding = new Thickness(4, 2, 4, 2);
			btnSearchDown.ToolTip = "Search next from top";
			btnSearchDown.Click += SearchDown_Click;
			sourceToolBar.Items.Add(btnSearchDown);

			Button btnSearchUp = new Button();
			StackPanel spUp = new StackPanel();
			spUp.Orientation = Orientation.Horizontal;

			TextBlock iconSearchUp = new TextBlock();
			iconSearchUp.Text = "\uE721";
			iconSearchUp.FontFamily = new FontFamily("Segoe Fluent Icons");
			iconSearchUp.FontSize = 16;
			iconSearchUp.VerticalAlignment = VerticalAlignment.Center;
			spUp.Children.Add(iconSearchUp);

			TextBlock iconArrowUp = new TextBlock();
			iconArrowUp.Text = "\uE761";
			iconArrowUp.FontFamily = new FontFamily("Segoe Fluent Icons");
			iconArrowUp.FontSize = 16;
			iconArrowUp.VerticalAlignment = VerticalAlignment.Center;
			iconArrowUp.RenderTransformOrigin = new Point(0.5, 0.5);
			iconArrowUp.RenderTransform = new RotateTransform(90);
			spUp.Children.Add(iconArrowUp);

			btnSearchUp.Content = spUp;
			btnSearchUp.Width = 48;
			btnSearchUp.Padding = new Thickness(4, 2, 4, 2);
            btnSearchUp.Margin = new Thickness(2, 0, 0, 0);
			btnSearchUp.ToolTip = "Search next from bottom";
			btnSearchUp.Click += SearchUp_Click;
			sourceToolBar.Items.Add(btnSearchUp);

			sourceToolBar.Items.Add(new Separator());

			Button btnFontSmaller = new Button();
			btnFontSmaller.Content = "\uE8E7";
			btnFontSmaller.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnFontSmaller.FontSize = 16;
			btnFontSmaller.Padding = new Thickness(6, 2, 6, 2);
            btnFontSmaller.Margin = new Thickness(2, 0, 0, 0);
			btnFontSmaller.ToolTip = "Decrease font size";
			btnFontSmaller.Click += (s, ev) =>
			{
			    if (sourceView.FontSize > 6)
			    {
			        sourceView.FontSize -= 1;
			        if (!string.IsNullOrEmpty(currentSource))
			            DisplaySource(currentSource);
			    }
			};
			sourceToolBar.Items.Add(btnFontSmaller);

			Button btnFontLarger = new Button();
			btnFontLarger.Content = "\uE8E8";
			btnFontLarger.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnFontLarger.FontSize = 16;
			btnFontLarger.Padding = new Thickness(6, 2, 6, 2);
            btnFontLarger.Margin = new Thickness(2, 0, 0, 0);
			btnFontLarger.ToolTip = "Increase font size";
			btnFontLarger.Click += (s, ev) =>
			{
			    if (sourceView.FontSize < 30)
			    {
			        sourceView.FontSize += 1;
			        if (!string.IsNullOrEmpty(currentSource))
			            DisplaySource(currentSource);
			    }
			};
			sourceToolBar.Items.Add(btnFontLarger);

			sourceToolBar.Items.Add(new Separator());

			findCount = new TextBlock();
			findCount.VerticalAlignment = VerticalAlignment.Center;
			findCount.Margin = new Thickness(4, 0, 0, 0);
			sourceToolBar.Items.Add(findCount);

			sourceToolBarTray.ToolBars.Add(sourceToolBar);
			sourcePanel.Children.Add(sourceToolBarTray);

			sourceView = new RichTextBox();
			sourceView.IsReadOnly = true;
			sourceView.FontFamily = new FontFamily("Consolas");
			sourceView.FontSize = windowSettings.SourceFontSize;
			sourceView.Foreground = new SolidColorBrush(Color.FromRgb(40, 80, 180));
			sourceView.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
			sourceView.AutoWordSelection = false;
			sourcePanel.Children.Add(sourceView);

			sourceTab.Content = sourcePanel;
			rightTabs.Items.Add(sourceTab);

            // Tab 2: App Settings
            TabItem appSettingsTab = new TabItem();
            appSettingsTab.Header = "⚙ App Settings";

            DockPanel appSettingsOuter = new DockPanel();

            // App Settings toolbar at the top
            ToolBarTray appSettingsToolBarTray = new ToolBarTray();
            DockPanel.SetDock(appSettingsToolBarTray, Dock.Top);

            ToolBar appSettingsToolBar = new ToolBar();
            appSettingsToolBar.Band = 0;
            appSettingsToolBar.BandIndex = 0;

            TextBlock appSettingsTitle = new TextBlock();
            appSettingsTitle.Text = "Application Settings";
            appSettingsTitle.FontWeight = FontWeights.Bold;
            appSettingsTitle.FontSize = 13;
            appSettingsTitle.VerticalAlignment = VerticalAlignment.Center;
            appSettingsTitle.Margin = new Thickness(0, 0, 8, 0);
            appSettingsToolBar.Items.Add(appSettingsTitle);

            appSettingsToolBar.Items.Add(new Separator());

            Button btnApplyColumns = new Button();
            btnApplyColumns.Content = "\xe930";
            btnApplyColumns.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnApplyColumns.FontSize = 16;
            btnApplyColumns.Padding = new Thickness(6, 2, 6, 2);
            btnApplyColumns.ToolTip = "Apply column visibility changes to the list";
            btnApplyColumns.Click += ApplyColumnSettings_Click;
            appSettingsToolBar.Items.Add(btnApplyColumns);

            Button btnSaveSettings = new Button();
            btnSaveSettings.Content = "\uE74E";
            btnSaveSettings.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnSaveSettings.FontSize = 16;
            btnSaveSettings.Padding = new Thickness(6, 2, 6, 2);
            btnSaveSettings.ToolTip = "Save all application settings now";
            btnSaveSettings.Click += SaveAppSettings_Click;
            appSettingsToolBar.Items.Add(btnSaveSettings);

            Button btnRestoreDefaults = new Button();
            btnRestoreDefaults.Content = "\xe845";
            btnRestoreDefaults.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnRestoreDefaults.FontSize = 16;
            btnRestoreDefaults.Padding = new Thickness(6, 2, 6, 2);
            btnRestoreDefaults.ToolTip = "Reset all column settings to defaults";
            btnRestoreDefaults.Click += RestoreDefaultColumns_Click;
            appSettingsToolBar.Items.Add(btnRestoreDefaults);

            appSettingsToolBarTray.ToolBars.Add(appSettingsToolBar);
            appSettingsOuter.Children.Add(appSettingsToolBarTray);

            // Scrollable content area below the toolbar
            ScrollViewer appSettingsScroll = new ScrollViewer();
            appSettingsScroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

            StackPanel appSettingsPanel = new StackPanel();
            appSettingsPanel.Margin = new Thickness(12);

            // Toolbar position
            TextBlock toolbarLabel = new TextBlock();
            toolbarLabel.Text = "Toolbar Position:";
            toolbarLabel.Margin = new Thickness(0, 0, 0, 4);
            appSettingsPanel.Children.Add(toolbarLabel);

            ComboBox toolbarCombo = new ComboBox();
            toolbarCombo.Width = 150;
            toolbarCombo.HorizontalAlignment = HorizontalAlignment.Left;
            toolbarCombo.Margin = new Thickness(0, 0, 0, 16);
            toolbarCombo.Items.Add("Top");
            toolbarCombo.Items.Add("Bottom");
            toolbarCombo.Items.Add("Left");
            toolbarCombo.Items.Add("Right");
            toolbarCombo.SelectedIndex = 0;
            toolbarCombo.SelectionChanged += (s, ev) =>
            {
                string pos = toolbarCombo.SelectedItem as string;
                if (pos != null)
                {
                    DockToolbarTo(pos);
                }
            };
            appSettingsPanel.Children.Add(toolbarCombo);

            // Editor path setting
            TextBlock editorLabel = new TextBlock();
            editorLabel.Text = "Text Editor Path:";
            editorLabel.Margin = new Thickness(0, 0, 0, 4);
            appSettingsPanel.Children.Add(editorLabel);

            TextBlock editorHint = new TextBlock();
            editorHint.Text = "Path to a text editor for opening .track files. Can be relative to app folder. Leave blank for system default.";
            editorHint.Foreground = Brushes.Gray;
            editorHint.FontSize = 11;
            editorHint.TextWrapping = TextWrapping.Wrap;
            editorHint.Margin = new Thickness(0, 0, 0, 4);
            appSettingsPanel.Children.Add(editorHint);

            DockPanel editorDock = new DockPanel();
            editorDock.Margin = new Thickness(0, 0, 0, 16);

            Button btnBrowseEditor = new Button();
            btnBrowseEditor.Content = "\xe838";
            btnBrowseEditor.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnBrowseEditor.FontSize = 16;
            btnBrowseEditor.Padding = new Thickness(8, 4, 8, 4);
            btnBrowseEditor.Margin = new Thickness(4, 0, 0, 0);
            btnBrowseEditor.ToolTip = "Browse for a text editor executable";
            btnBrowseEditor.Click += BrowseEditor_Click;
            DockPanel.SetDock(btnBrowseEditor, Dock.Right);
            editorDock.Children.Add(btnBrowseEditor);

            editEditorPath = new TextBox();
            editEditorPath.Text = windowSettings.EditorPath ?? "";
            editorDock.Children.Add(editEditorPath);

            appSettingsPanel.Children.Add(editorDock);

            // Column Visibility section
            TextBlock colTitle = new TextBlock();
            colTitle.Text = "Column Visibility:";
            colTitle.FontWeight = FontWeights.Bold;
            colTitle.Margin = new Thickness(0, 0, 0, 4);
            appSettingsPanel.Children.Add(colTitle);

            TextBlock colHint = new TextBlock();
            colHint.Text = "Show/hide columns. Drag column headers to reorder. Drag column edges to resize.";
            colHint.Foreground = Brushes.Gray;
            colHint.FontSize = 11;
            colHint.TextWrapping = TextWrapping.Wrap;
            colHint.Margin = new Thickness(0, 0, 0, 8);
            appSettingsPanel.Children.Add(colHint);

            // Build checkbox list from current or default settings
            List<ColumnSetting> currentColSettings = windowSettings.ColumnSettings;
            if (currentColSettings.Count == 0)
                currentColSettings = WindowSettings.GetDefaultColumns();
            currentColSettings.Sort((a, b) => a.DisplayIndex.CompareTo(b.DisplayIndex));

            columnCheckboxPanel = new StackPanel();
            columnCheckboxPanel.Margin = new Thickness(0, 0, 0, 16);

            foreach (ColumnSetting cs in currentColSettings)
            {
                CheckBox chk = new CheckBox();
                chk.Content = cs.Header + "  (" + cs.Binding + ")";
                chk.IsChecked = cs.Visible;
                chk.Tag = cs.Binding;
                chk.Margin = new Thickness(0, 2, 0, 2);
                columnCheckboxPanel.Children.Add(chk);
            }

            appSettingsPanel.Children.Add(columnCheckboxPanel);

            appSettingsScroll.Content = appSettingsPanel;
            appSettingsOuter.Children.Add(appSettingsScroll);
            appSettingsTab.Content = appSettingsOuter;

            // Tab 3: WebView2
            TabItem webTab = new TabItem();
            webTab.Header = "🌐 WebView";

            DockPanel webPanel = new DockPanel();

            // ---- TOP TOOLBAR: Search engine + address bar + nav buttons ----
            ToolBarTray webTopToolBarTray = new ToolBarTray();
            DockPanel.SetDock(webTopToolBarTray, Dock.Top);

            ToolBar webTopToolBar = new ToolBar();
            webTopToolBar.Band = 0;
            webTopToolBar.BandIndex = 0;

            ComboBox searchEngineCombo = new ComboBox();
            searchEngineCombo.Width = 110;
            searchEngineCombo.Items.Add("DuckDuckGo");
            searchEngineCombo.Items.Add("Google");
            searchEngineCombo.Items.Add("Bing");
            searchEngineCombo.Items.Add("Startpage");
            searchEngineCombo.Items.Add("Brave");
            searchEngineCombo.Items.Add("Yahoo");
            searchEngineCombo.SelectedIndex = 0;
            webTopToolBar.Items.Add(searchEngineCombo);

            TextBox webSearchBox = new TextBox();
            webSearchBox.Width = 300;
            webSearchBox.ToolTip = "Enter a URL or search term, then press Enter or click the search button";
            webTopToolBar.Items.Add(webSearchBox);

            Button btnWebGo = new Button();
            btnWebGo.Content = "\uE721";
            btnWebGo.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnWebGo.FontSize = 16;
            btnWebGo.Padding = new Thickness(6, 2, 6, 2);
            btnWebGo.Margin = new Thickness(4, 0, 0, 0);
            btnWebGo.ToolTip = "Search or navigate to URL";
            webTopToolBar.Items.Add(btnWebGo);

            webTopToolBar.Items.Add(new Separator());

            Button btnWebBack = new Button();
            btnWebBack.Content = "\xe760";
            btnWebBack.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnWebBack.FontSize = 16;
            btnWebBack.Padding = new Thickness(6, 2, 6, 2);
            btnWebBack.ToolTip = "Back";
            btnWebBack.Click += (s, ev) =>
            {
                try { if (webView.CoreWebView2 != null && webView.CanGoBack) webView.GoBack(); }
                catch (Exception) { }
            };
            webTopToolBar.Items.Add(btnWebBack);

            Button btnWebForward = new Button();
            btnWebForward.Content = "\xe761";
            btnWebForward.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnWebForward.FontSize = 16;
            btnWebForward.Padding = new Thickness(6, 2, 6, 2);
            btnWebForward.Margin = new Thickness(2, 0, 0, 0);
            btnWebForward.ToolTip = "Forward";
            btnWebForward.Click += (s, ev) =>
            {
                try { if (webView.CoreWebView2 != null && webView.CanGoForward) webView.GoForward(); }
                catch (Exception) { }
            };
            webTopToolBar.Items.Add(btnWebForward);

            Button btnWebHome = new Button();
            btnWebHome.Content = "\xe80f";
            btnWebHome.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnWebHome.FontSize = 16;
            btnWebHome.Padding = new Thickness(6, 2, 6, 2);
            btnWebHome.Margin = new Thickness(2, 0, 0, 0);
            btnWebHome.ToolTip = "Return to current track URL";
            btnWebHome.Click += (s, ev) =>
            {
                try
                {
                    if (webView.CoreWebView2 != null && currentTrackItem != null &&
                        !string.IsNullOrWhiteSpace(currentTrackItem.TrackURL))
                    {
                        string url = currentTrackItem.TrackURL;
                        if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                            url = "https://" + url;
                        webView.CoreWebView2.Navigate(url);
                    }
                }
                catch (Exception) { }
            };
            webTopToolBar.Items.Add(btnWebHome);

            webTopToolBarTray.ToolBars.Add(webTopToolBar);

            // Search/navigate action
            Action doWebSearch = () =>
            {
                string input = webSearchBox.Text.Trim();
                if (string.IsNullOrEmpty(input)) return;

                string navigateUrl;

                if (input.StartsWith("http://") || input.StartsWith("https://") ||
                    (input.Contains(".") && !input.Contains(" ")))
                {
                    navigateUrl = input;
                    if (!navigateUrl.StartsWith("http://") && !navigateUrl.StartsWith("https://"))
                        navigateUrl = "https://" + navigateUrl;
                }
                else
                {
                    string query = Uri.EscapeDataString(input);
                    string engine = searchEngineCombo.SelectedItem as string ?? "DuckDuckGo";

                    switch (engine)
                    {
                        case "Google":
                            navigateUrl = "https://www.google.com/search?q=" + query;
                            break;
                        case "Bing":
                            navigateUrl = "https://www.bing.com/search?q=" + query;
                            break;
                        case "Startpage":
                            navigateUrl = "https://www.startpage.com/do/search?q=" + query;
                            break;
                        case "Brave":
                            navigateUrl = "https://search.brave.com/search?q=" + query;
                            break;
                        case "Yahoo":
                            navigateUrl = "https://search.yahoo.com/search?p=" + query;
                            break;
                        default:
                            navigateUrl = "https://duckduckgo.com/?q=" + query;
                            break;
                    }
                }

                try
                {
                    if (webView.CoreWebView2 != null)
                    {
                        webView.CoreWebView2.Navigate(navigateUrl);
                        statusFile.Text = "Navigating: " + navigateUrl;
                    }
                }
                catch (Exception ex)
                {
                    statusFile.Text = "Navigation error: " + ex.Message;
                }
            };

            btnWebGo.Click += (s, ev) => doWebSearch();
            webSearchBox.KeyDown += (s, ev) =>
            {
                if (ev.Key == System.Windows.Input.Key.Enter)
                    doWebSearch();
            };

            // ---- BOTTOM TOOLBAR: Zoom, cookies, clear data, checkbox ----
            ToolBarTray webBottomToolBarTray = new ToolBarTray();
            DockPanel.SetDock(webBottomToolBarTray, Dock.Bottom);

            ToolBar webBottomToolBar = new ToolBar();
            webBottomToolBar.Band = 0;
            webBottomToolBar.BandIndex = 0;

            Button btnZoomIn = new Button();
            btnZoomIn.Content = "\xe8a3";
            btnZoomIn.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnZoomIn.FontSize = 16;
            btnZoomIn.Padding = new Thickness(6, 2, 6, 2);
            btnZoomIn.ToolTip = "Zoom in";
            btnZoomIn.Click += ZoomIn_Click;
            webBottomToolBar.Items.Add(btnZoomIn);

            Button btnZoomOut = new Button();
            btnZoomOut.Content = "\xe71f";
            btnZoomOut.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnZoomOut.FontSize = 16;
            btnZoomOut.Padding = new Thickness(6, 2, 6, 2);
            btnZoomOut.Margin = new Thickness(2, 0, 0, 0);
            btnZoomOut.ToolTip = "Zoom out";
            btnZoomOut.Click += ZoomOut_Click;
            webBottomToolBar.Items.Add(btnZoomOut);

            Button btnZoomReset = new Button();
            btnZoomReset.Content = "100%";
            btnZoomReset.Padding = new Thickness(6, 2, 6, 2);
            btnZoomReset.Margin = new Thickness(2, 0, 0, 0);
            btnZoomReset.ToolTip = "Reset zoom";
            btnZoomReset.Click += ZoomReset_Click;
            webBottomToolBar.Items.Add(btnZoomReset);

            Button btnZoomFit = new Button();
            btnZoomFit.Content = "\xe9a6";
            btnZoomFit.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnZoomFit.FontSize = 16;
            btnZoomFit.Padding = new Thickness(6, 2, 6, 2);
            btnZoomFit.Margin = new Thickness(2, 0, 0, 0);
            btnZoomFit.ToolTip = "Fit to panel width";
            btnZoomFit.Click += ZoomFit_Click;
            webBottomToolBar.Items.Add(btnZoomFit);

            webBottomToolBar.Items.Add(new Separator());

            Button btnClearCookies = new Button();
            btnClearCookies.Content = "🍪";
            btnClearCookies.FontSize = 16;
            btnClearCookies.Padding = new Thickness(6, 2, 6, 2);
            btnClearCookies.ToolTip = "Delete all cookies";
            btnClearCookies.Click += ClearCookies_Click;
            webBottomToolBar.Items.Add(btnClearCookies);

            Button btnClearAll = new Button();
            btnClearAll.Content = "\xecad";
            btnClearAll.FontFamily = new FontFamily("Segoe Fluent Icons");
            btnClearAll.FontSize = 16;
            btnClearAll.Padding = new Thickness(6, 2, 6, 2);
            btnClearAll.Margin = new Thickness(2, 0, 0, 0);
            btnClearAll.ToolTip = "Clear all cookies, cache, and browsing data";
            btnClearAll.Click += ClearAllData_Click;
            webBottomToolBar.Items.Add(btnClearAll);

            webBottomToolBar.Items.Add(new Separator());

            CheckBox chkBlockPopups = new CheckBox();
            chkBlockPopups.Content = "Block Cookie Popups";
            chkBlockPopups.VerticalAlignment = VerticalAlignment.Center;
            chkBlockPopups.IsChecked = true;
            chkBlockPopups.ToolTip = "Automatically hides cookie consent popups.\nUnchecking clears cookies for this site\n(you may be signed out) and reload the page.";
            chkBlockPopups.Checked += async (s, ev) =>
            {
                blockCookiePopups = true;
                statusFile.Text = "Cookie popup blocker enabled";
                try
                {
                    if (webView.CoreWebView2 != null)
                        await InjectCookiePopupBlocker();
                }
                catch (Exception) { }
            };
            chkBlockPopups.Unchecked += async (s, ev) =>
            {
                blockCookiePopups = false;
                statusFile.Text = "Cookie popup blocker disabled — clearing site cookies and reloading...";
                try
                {
                    if (webView.CoreWebView2 != null)
                    {
                        await webView.CoreWebView2.ExecuteScriptAsync(@"
                            window.__patCookieBlockerActive = false;
                            var css = document.getElementById('__patCookieBlockerCSS');
                            if (css) css.remove();
                        ");

                        string currentUrl = webView.CoreWebView2.Source;

                        if (!string.IsNullOrEmpty(currentUrl) && currentUrl != "about:blank")
                        {
                            try
                            {
                                var cookies = await webView.CoreWebView2.CookieManager.GetCookiesAsync(currentUrl);
                                foreach (var cookie in cookies)
                                {
                                    webView.CoreWebView2.CookieManager.DeleteCookie(cookie);
                                }
                            }
                            catch (Exception) { }

                            await webView.CoreWebView2.ExecuteScriptAsync(@"
                                try { localStorage.clear(); } catch(e) {}
                                try { sessionStorage.clear(); } catch(e) {}
                            ");

                            webView.CoreWebView2.Navigate(currentUrl);
                        }
                    }
                }
                catch (Exception) { }
            };
            webBottomToolBar.Items.Add(chkBlockPopups);

            webBottomToolBarTray.ToolBars.Add(webBottomToolBar);

            // ---- Add to DockPanel in correct order ----
            // Top toolbar first (docked top)
            webPanel.Children.Add(webTopToolBarTray);
            // Bottom toolbar second (docked bottom)
            webPanel.Children.Add(webBottomToolBarTray);

            // WebView2 last — it fills the remaining space
            webView = new Microsoft.Web.WebView2.Wpf.WebView2();
            webPanel.Children.Add(webView);

            webTab.Content = webPanel;
			rightTabs.Items.Add(webTab);
			rightTabs.Items.Add(BuildThemeTab());
			rightTabs.Items.Add(appSettingsTab);
        }

        // ============================
        // Settings Field Helpers
        // ============================

        private TextBox AddSettingsField(StackPanel parent, string label, double width = 0)
        {
            TextBlock lbl = new TextBlock();
            lbl.Text = label;
            lbl.Margin = new Thickness(0, 0, 0, 4);
            parent.Children.Add(lbl);

            TextBox txt = new TextBox();
            txt.Margin = new Thickness(0, 0, 0, 8);
            if (width > 0)
            {
                txt.Width = width;
                txt.HorizontalAlignment = HorizontalAlignment.Left;
            }
            parent.Children.Add(txt);
            return txt;
        }

        private void ApplyCharacterSelection(TextBox txt)
        {
            int selectionAnchor = -1;

            txt.PreviewMouseLeftButtonDown += (s, ev) =>
            {
                TextBox tb = s as TextBox;
                if (tb == null) return;

                int clickPos = tb.GetCharacterIndexFromPoint(ev.GetPosition(tb), true);
                if (clickPos >= 0)
                {
                    selectionAnchor = clickPos;
                    tb.Focus();
                    tb.CaretIndex = clickPos;
                    tb.Select(clickPos, 0);
                    ev.Handled = true;
                    tb.CaptureMouse();
                }
            };

            txt.PreviewMouseMove += (s, ev) =>
            {
                TextBox tb = s as TextBox;
                if (tb == null) return;

                if (ev.LeftButton == System.Windows.Input.MouseButtonState.Pressed
                    && selectionAnchor >= 0 && tb.IsMouseCaptured)
                {
                    int currentPos = tb.GetCharacterIndexFromPoint(ev.GetPosition(tb), true);
                    if (currentPos >= 0)
                    {
                        int start = Math.Min(selectionAnchor, currentPos);
                        int length = Math.Abs(currentPos - selectionAnchor);
                        tb.Select(start, length);
                    }
                    ev.Handled = true;
                }
            };

            txt.PreviewMouseLeftButtonUp += (s, ev) =>
            {
                TextBox tb = s as TextBox;
                if (tb != null && tb.IsMouseCaptured)
                {
                    tb.ReleaseMouseCapture();
                }
                selectionAnchor = -1;
            };
        }

        private TextBox AddSettingsMultiline(StackPanel parent, string label)
        {
            TextBlock lbl = new TextBlock();
            lbl.Text = label;
            lbl.Margin = new Thickness(0, 0, 0, 4);
            parent.Children.Add(lbl);

            TextBox txt = new TextBox();
            txt.Height = 60;
            txt.AcceptsReturn = true;
            txt.TextWrapping = TextWrapping.Wrap;
            txt.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;
            txt.Margin = new Thickness(0, 0, 0, 8);
            ApplyCharacterSelection(txt);
            parent.Children.Add(txt);
            return txt;
        }

		private static readonly StatusToColorConverter statusConverter = new StatusToColorConverter();

		private GridViewColumn MakeColumn(string header, string binding, double width)
		{
		    GridViewColumn col = new GridViewColumn();
		    col.Header = header;
		    col.Width = width;

		    DataTemplate template = new DataTemplate();
		    FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
		    textBlock.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(binding));
		    textBlock.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
		    textBlock.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
		    textBlock.SetValue(TextBlock.MaxHeightProperty, 20.0);
		    template.VisualTree = textBlock;
		    col.CellTemplate = template;

		    return col;
		}

		private GridViewColumn MakeColumn(string header, string binding, double width, string statusBinding = null)
		{
		    GridViewColumn col = new GridViewColumn();
		    col.Header = header;
		    col.Width = width;

		    DataTemplate template = new DataTemplate();

		    FrameworkElementFactory border = new FrameworkElementFactory(typeof(Border));
		    border.SetValue(Border.PaddingProperty, new Thickness(2, 0, 2, 0));
		    border.SetValue(Border.MarginProperty, new Thickness(-6, -2, -6, -2));

		    FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
		    textBlock.SetBinding(TextBlock.TextProperty, new System.Windows.Data.Binding(binding));
		    textBlock.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis);
		    textBlock.SetValue(TextBlock.TextWrappingProperty, TextWrapping.NoWrap);
		    textBlock.SetValue(TextBlock.MaxHeightProperty, 20.0);

		    // Only apply status coloring if a statusBinding was provided
		    if (statusBinding != null)
		    {
		        System.Windows.Data.Binding fgBinding = new System.Windows.Data.Binding(statusBinding);
		        fgBinding.Converter = new StatusToColorConverter();
		        textBlock.SetBinding(TextBlock.ForegroundProperty, fgBinding);
		    }

		    border.AppendChild(textBlock);
		    template.VisualTree = border;
		    col.CellTemplate = template;

		    return col;
		}
		
        private string GetStatusBindingForColumn(string binding)
        {
            switch (binding)
            {
                case "TrackURL": return "TrackURLStatus";
                case "StartString": return "StartStringStatus";
                case "StopString": return "StopStringStatus";
                case "TrackBlockHash": return "HashStatus";
                case "LatestVersion": return "LatestVersionStatus";
                case "DownloadURL": return "DownloadURLStatus";
                case "DownloadSizeKb": return "DownloadSizeStatus";
                default: return null;
            }
        }

        // ============================
        // Layout Placement
        // ============================

        private void PlaceControlsInLayout()
        {
            RemoveFromParent(categoryPanel);
            RemoveFromParent(itemList);
            RemoveFromParent(trackSettingsPanel);
            RemoveFromParent(rightTabs);

            if (isHorizontalLayout)
            {
                catHostH.Child = categoryPanel;
                listHostH.Child = itemList;
                trackSettingsHostH.Child = trackSettingsPanel;
                tabsHostH.Child = rightTabs;
            }
            else
            {
                catHostV.Child = categoryPanel;
                listHostV.Child = itemList;
                trackSettingsHostV.Child = trackSettingsPanel;
                tabsHostV.Child = rightTabs;
            }
        }

        private void RemoveFromParent(UIElement element)
        {
            if (catHostV.Child == element) catHostV.Child = null;
            if (listHostV.Child == element) listHostV.Child = null;
            if (trackSettingsHostV.Child == element) trackSettingsHostV.Child = null;
            if (tabsHostV.Child == element) tabsHostV.Child = null;
            if (catHostH.Child == element) catHostH.Child = null;
            if (listHostH.Child == element) listHostH.Child = null;
            if (trackSettingsHostH.Child == element) trackSettingsHostH.Child = null;
            if (tabsHostH.Child == element) tabsHostH.Child = null;
        }

        // ============================
        // Category Tree
        // ============================

        private void RefreshCategoryTree()
        {
            categoryTree.Items.Clear();

            if (!Directory.Exists(categoriesPath))
                return;

            foreach (string dir in Directory.GetDirectories(categoriesPath))
            {
                string folderName = Path.GetFileName(dir);
                int trackCount = Directory.GetFiles(dir, "*.track").Length;
                string sourceLinkFile = Path.Combine(dir, "_source.link");
                bool isLinked = File.Exists(sourceLinkFile);

                string header;
                if (isLinked)
                {
                    header = trackCount > 0
                        ? "📂 " + folderName + " (" + trackCount + ") [linked]"
                        : "📂 " + folderName + " [linked]";
                }
                else
                {
                    header = trackCount > 0
                        ? "📁 " + folderName + " (" + trackCount + ")"
                        : "📁 " + folderName;
                }

                TreeViewItem item = new TreeViewItem();
                item.Header = header;
                item.Tag = dir;
                item.IsExpanded = false;
                categoryTree.Items.Add(item);
            }

            statusFile.Text = categoryTree.Items.Count + " categories";
        }

        private void AddCategory_Click(object sender, RoutedEventArgs e)
        {
            string newName = ShowInputDialog("New Category", "Enter category name:");

            if (string.IsNullOrWhiteSpace(newName))
                return;

            foreach (char c in Path.GetInvalidFileNameChars())
            {
                newName = newName.Replace(c.ToString(), "");
            }

            if (string.IsNullOrWhiteSpace(newName))
            {
                MessageBox.Show("Invalid folder name.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string newPath = Path.Combine(categoriesPath, newName);

            if (Directory.Exists(newPath))
            {
                MessageBox.Show("Category '" + newName + "' already exists.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                Directory.CreateDirectory(newPath);
                RefreshCategoryTree();

                foreach (TreeViewItem item in categoryTree.Items)
                {
                    if ((string)item.Tag == newPath)
                    {
                        item.IsSelected = true;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating category: " + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveCategory_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem selected = categoryTree.SelectedItem as TreeViewItem;

            if (selected == null)
            {
                MessageBox.Show("No category selected.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string folderPath = (string)selected.Tag;
            string folderName = Path.GetFileName(folderPath);

            int fileCount = 0;
            if (Directory.Exists(folderPath))
            {
                fileCount = Directory.GetFiles(folderPath).Length;
            }

            string message = "Delete category '" + folderName + "'?";
            if (fileCount > 0)
            {
                message += "\n\nThis folder contains " + fileCount +
                    " file(s) which will also be deleted!";
            }

            MessageBoxResult result = MessageBox.Show(message, "Delete Category",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    Directory.Delete(folderPath, true);
                    RefreshCategoryTree();
                    currentItems.Clear();
                    itemList.ItemsSource = null;
                    ClearTrackFields();
                    statusFile.Text = "Category deleted: " + folderName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting category: " + ex.Message,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void OpenCategoryFolder_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem selected = categoryTree.SelectedItem as TreeViewItem;

            if (selected != null)
            {
                string folderPath = (string)selected.Tag;
                if (Directory.Exists(folderPath))
                {
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(folderPath)
                        { UseShellExecute = true });
                    return;
                }
            }

            // No category selected — open the root Categories folder
            if (Directory.Exists(categoriesPath))
            {
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo(categoriesPath)
                    { UseShellExecute = true });
            }
        }

        private void CategoryTree_Selected(object sender,
            RoutedPropertyChangedEventArgs<object> e)
        {
            TreeViewItem selected = e.NewValue as TreeViewItem;

            if (selected == null)
                return;

            string folderPath = (string)selected.Tag;
            currentCategoryPath = folderPath;
            LoadTrackFiles(folderPath);
        }

        // ============================
        // Track File Loading
        // ============================

        private void LoadTrackFiles(string folderPath)
        {
            currentItems.Clear();
            string folderName = Path.GetFileName(folderPath);

            if (!Directory.Exists(folderPath))
                return;

            string[] trackFiles = Directory.GetFiles(folderPath, "*.track");

            foreach (string file in trackFiles)
            {
                TrackItem item = TrackItem.LoadFromFile(file);
                currentItems.Add(item);
            }

            itemList.ItemsSource = null;
            itemList.ItemsSource = currentItems;

            if (currentItems.Count > 0)
            {
                itemList.SelectedIndex = 0;
            }
            else
            {
                ClearTrackFields();
            }

            statusFile.Text = "Category: " + folderName + " — " + currentItems.Count + " items";
        }

        // ============================
        // Item List
        // ============================

        private void ItemList_SelectionChanged(object sender,
            SelectionChangedEventArgs e)
        {
            TrackItem selected = itemList.SelectedItem as TrackItem;

            if (selected == null)
                return;

            currentTrackItem = selected;

            editName.Text = selected.ProgramName;
            editTrackURL.Text = selected.TrackURL;
            editStartString.Text = selected.StartString;
            editStopString.Text = selected.StopString;
            editDownloadURL.Text = selected.DownloadURL;
            editVersion.Text = selected.Version;
            UpdateVersionDisplay();
            editReleaseDate.Text = selected.ReleaseDate;
            editPublisherName.Text = selected.PublisherName;
            editSuiteName.Text = selected.SuiteName;

            statusFile.Text = "Track: " + selected.ProgramName +
                " (" + Path.GetFileName(selected.FilePath) + ")";

            // Auto-download source if track has a URL
            if (!suppressAutoDownload && !string.IsNullOrWhiteSpace(selected.TrackURL))
            {
                AutoDownloadSource(selected.TrackURL);
            }
            else if (!suppressAutoDownload)
            {
                currentSource = "";
                sourceView.Document.Blocks.Clear();
            }
            // If suppressAutoDownload is true, leave currentSource and sourceView untouched
        }
        
        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            CheckBox cb = sender as CheckBox;
            bool isChecked = cb?.IsChecked ?? false;

            foreach (TrackItem item in currentItems)
            {
                item.IsSelected = isChecked;
            }

            itemList.ItemsSource = null;
            itemList.ItemsSource = currentItems;
        }

        private void ColumnHeader_Click(object sender, RoutedEventArgs e)
        {
            GridViewColumnHeader header = e.OriginalSource as GridViewColumnHeader;

            if (header == null || header.Column == null)
                return;

            string sortBy = "";
            if (header.Column.DisplayMemberBinding is System.Windows.Data.Binding binding)
            {
                sortBy = binding.Path.Path;
            }
            else
            {
                return;
            }

            if (string.IsNullOrEmpty(sortBy))
                return;

            if (sortBy == lastSortColumn)
            {
                lastSortAscending = !lastSortAscending;
            }
            else
            {
                lastSortColumn = sortBy;
                lastSortAscending = true;
            }

            currentItems.Sort((a, b) =>
            {
                string valA = GetPropertyValue(a, sortBy);
                string valB = GetPropertyValue(b, sortBy);
                int result = string.Compare(valA, valB,
                    StringComparison.OrdinalIgnoreCase);
                return lastSortAscending ? result : -result;
            });

            itemList.ItemsSource = null;
            itemList.ItemsSource = currentItems;
        }

        private string GetPropertyValue(TrackItem item, string propertyName)
        {
            var prop = typeof(TrackItem).GetProperty(propertyName);
            if (prop != null)
            {
                object val = prop.GetValue(item);
                return val?.ToString() ?? "";
            }
            return "";
        }

        private async void CheckSelected_Click(object sender, RoutedEventArgs e)
        {
            List<TrackItem> toCheck = new List<TrackItem>();

            foreach (TrackItem item in currentItems)
            {
                if (item.IsSelected)
                {
                    toCheck.Add(item);
                }
            }

            if (toCheck.Count == 0)
            {
                MessageBox.Show("No tracks selected. Use the checkboxes to select tracks.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await BatchCheck(toCheck);
        }

        private async void CheckAll_Click(object sender, RoutedEventArgs e)
        {
            if (currentItems.Count == 0)
            {
                MessageBox.Show("No tracks in current category.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            await BatchCheck(currentItems);
        }

        private async System.Threading.Tasks.Task BatchCheck(List<TrackItem> items)
        {
            int total = items.Count;
            int current = 0;
            int changed = 0;
            int unchanged = 0;
            int errors = 0;
            int newChecks = 0;

            statusProgress.IsIndeterminate = false;
            statusProgress.Minimum = 0;
            statusProgress.Maximum = total;
            statusProgress.Value = 0;

            btnCheckSelected.IsEnabled = false;
            btnCheckAll.IsEnabled = false;

            try
            {
                foreach (TrackItem item in items)
                {
                    current++;
                    statusFile.Text = "Checking " + current + "/" + total +
                        ": " + item.ProgramName;
                    statusProgress.Value = current;

                    if (string.IsNullOrWhiteSpace(item.TrackURL))
                    {
                        item.TrackStatus = "error";
                        item.LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                        errors++;
                        continue;
                    }

                    string url = item.TrackURL;
                    if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                    {
                        url = "https://" + url;
                    }

                    try
                    {
		                SetupHttpHeaders(url);

                        string source = await httpClient.GetStringAsync(url);
                        long downloadBytes = System.Text.Encoding.UTF8.GetByteCount(source);

                        CheckResult result = item.ApplyCheck(source, downloadBytes);
                        item.SaveToFile();

                        switch (result.Status)
                        {
                            case "changed": changed++; break;
                            case "unchanged": unchanged++; break;
                            case "new": newChecks++; break;
                            case "error": errors++; break;
                        }
                    }
                    catch (Exception)
                    {
                        // HttpClient failed — try WebView2 fallback
                        try
                        {
                            statusFile.Text = "Checking " + current + "/" + total +
                                ": " + item.ProgramName + " (WebView2 fallback)";

                            string source = await DownloadViaWebView(url);
                            if (!string.IsNullOrEmpty(source))
                            {
                                long downloadBytes = System.Text.Encoding.UTF8.GetByteCount(source);
                                CheckResult result = item.ApplyCheck(source, downloadBytes);
                                item.SaveToFile();

                                switch (result.Status)
                                {
                                    case "changed": changed++; break;
                                    case "unchanged": unchanged++; break;
                                    case "new": newChecks++; break;
                                    case "error": errors++; break;
                                }
                            }
                            else
                            {
                                item.TrackStatus = "error";
                                item.TrackURLStatus = "error";
                                item.StartStringStatus = "";
                                item.StopStringStatus = "";
                                item.HashStatus = "";
                                item.LatestVersionStatus = "";
                                item.LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                                item.SaveToFile();
                                errors++;
                            }
                        }
                        catch (Exception)
                        {
                            item.TrackStatus = "error";
                            item.TrackURLStatus = "error";
                            item.StartStringStatus = "";
                            item.StopStringStatus = "";
                            item.HashStatus = "";
                            item.LatestVersionStatus = "";
                            item.LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                            item.SaveToFile();
                            errors++;
                        }
                    }

                    // Small delay to avoid hammering servers
                    await System.Threading.Tasks.Task.Delay(500);
                }
            }
            finally
            {
                btnCheckSelected.IsEnabled = true;
                btnCheckAll.IsEnabled = true;
                statusProgress.IsIndeterminate = false;
                statusProgress.Value = 0;
                statusProgress.Maximum = 100;
            }

            // Refresh the list and preserve selection
            int prevIndex = currentItems.IndexOf(currentTrackItem);
            suppressAutoDownload = true;
            itemList.ItemsSource = null;
            itemList.ItemsSource = currentItems;
            if (prevIndex >= 0)
                itemList.SelectedIndex = prevIndex;
            suppressAutoDownload = false;

            // Update fields if a track is loaded
            if (currentTrackItem != null)
            {
                editVersion.Text = currentTrackItem.Version;
                UpdateVersionDisplay();
                editReleaseDate.Text = currentTrackItem.ReleaseDate;
            }

            // Re-download source for the currently selected track
            if (currentTrackItem != null && !string.IsNullOrWhiteSpace(currentTrackItem.TrackURL))
            {
                AutoDownloadSource(currentTrackItem.TrackURL);
            }
        }

        private DataTemplate CreateStatusCellTemplate()
        {
            DataTemplate template = new DataTemplate();

            FrameworkElementFactory textBlock = new FrameworkElementFactory(typeof(TextBlock));
            textBlock.SetValue(TextBlock.TextProperty, "\uF136"); // Filled circle
            textBlock.SetValue(TextBlock.FontFamilyProperty, new FontFamily("Segoe Fluent Icons"));
            textBlock.SetValue(TextBlock.FontSizeProperty, 14.0);
            textBlock.SetValue(TextBlock.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            textBlock.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);
            textBlock.AddHandler(TextBlock.LoadedEvent, new RoutedEventHandler(StatusIcon_Loaded));
            textBlock.SetBinding(TextBlock.TagProperty, new System.Windows.Data.Binding("TrackStatus"));

            template.VisualTree = textBlock;
            return template;
        }

        private void StatusIcon_Loaded(object sender, RoutedEventArgs e)
        {
            TextBlock tb = sender as TextBlock;
            if (tb == null) return;

            string status = tb.Tag as string;
            ApplyStatusColor(tb, status);

            // Also update when the data context changes
            tb.DataContextChanged += (s, ev) =>
            {
                TrackItem item = tb.DataContext as TrackItem;
                if (item != null)
                {
                    tb.Tag = item.TrackStatus;
                    ApplyStatusColor(tb, item.TrackStatus);
                }
            };
        }

		private void ApplyStatusColor(TextBlock tb, string status)
		{
		    switch (status)
		    {
		        case "unchanged":
		            tb.Foreground = new SolidColorBrush(currentTheme.StatusUnchanged);
		            tb.ToolTip = "Unchanged — no changes detected";
		            break;
		        case "changed":
		            tb.Foreground = new SolidColorBrush(currentTheme.StatusChanged);
		            tb.ToolTip = "Changed — update detected!";
		            break;
		        case "error":
		            tb.Foreground = new SolidColorBrush(currentTheme.StatusError);
		            tb.ToolTip = "Error — check failed";
		            break;
		        case "new":
		            tb.Foreground = new SolidColorBrush(currentTheme.StatusNew);
		            tb.ToolTip = "New — first check recorded";
		            break;
		        default:
		            tb.Foreground = new SolidColorBrush(currentTheme.StatusUnchecked);
		            tb.ToolTip = "Unchecked";
		            break;
		    }
		}

		private void SetBatchStatusText(int total, int changed, int unchanged, int newChecks, int errors)
		{
		    statusFile.Inlines.Clear();
		    statusFile.Inlines.Add(new Run("Batch complete: " + total + " checked — "));

		    Run changedIcon = new Run("\uF136 ");
		    changedIcon.FontFamily = new FontFamily("Segoe Fluent Icons");
		    changedIcon.Foreground = new SolidColorBrush(currentTheme.StatusChanged);
		    statusFile.Inlines.Add(changedIcon);
		    statusFile.Inlines.Add(new Run(changed + " changed  "));

		    Run unchangedIcon = new Run("\uF136 ");
		    unchangedIcon.FontFamily = new FontFamily("Segoe Fluent Icons");
		    unchangedIcon.Foreground = new SolidColorBrush(currentTheme.StatusUnchanged);
		    statusFile.Inlines.Add(unchangedIcon);
		    statusFile.Inlines.Add(new Run(unchanged + " unchanged  "));

		    Run newIcon = new Run("\uF136 ");
		    newIcon.FontFamily = new FontFamily("Segoe Fluent Icons");
		    newIcon.Foreground = new SolidColorBrush(currentTheme.StatusNew);
		    statusFile.Inlines.Add(newIcon);
		    statusFile.Inlines.Add(new Run(newChecks + " new  "));

		    Run errorIcon = new Run("\uF136 ");
		    errorIcon.FontFamily = new FontFamily("Segoe Fluent Icons");
		    errorIcon.Foreground = new SolidColorBrush(currentTheme.StatusError);
		    statusFile.Inlines.Add(errorIcon);
		    statusFile.Inlines.Add(new Run(errors + " errors"));
		}
		
        private async void AutoDownloadSource(string url)
        {
            if (!url.StartsWith("http://") && !url.StartsWith("https://"))
                url = "https://" + url;

            try
            {
                SetupHttpHeaders(url);
                string source = await httpClient.GetStringAsync(url);
                currentSource = source;
                DisplaySource(currentSource);

                // Navigate WebView too
                try
                {
                    await webView.EnsureCoreWebView2Async();
                    webView.CoreWebView2.Navigate(url);
                }
                catch (Exception) { }
            }
            catch (Exception)
            {
                // HttpClient failed — try WebView2 fallback (handles Cloudflare etc.)
                try
                {
                    statusFile.Text = "HttpClient blocked, trying WebView2 fallback...";
                    string source = await DownloadViaWebView(url);
                    if (!string.IsNullOrEmpty(source))
                    {
                        currentSource = source;
                        DisplaySource(currentSource);
                        statusFile.Text = "Loaded via WebView2: " + url;
                    }
                    else
                    {
                        currentSource = "";
                        sourceView.Document.Blocks.Clear();
                        statusFile.Text = "Failed to load: " + url;
                    }
                }
                catch (Exception)
                {
                    currentSource = "";
                    sourceView.Document.Blocks.Clear();
                }
            }
        }

        // ============================
        // Track Settings Buttons
        // ============================

        private void OpenFileInEditor_Click(object sender, RoutedEventArgs e)
        {
            TrackItem selected = itemList.SelectedItem as TrackItem;

            if (selected == null)
            {
                MessageBox.Show("No track selected.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!File.Exists(selected.FilePath))
            {
                MessageBox.Show("Track file not found:\n" + selected.FilePath,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string editorPath = editEditorPath.Text.Trim();

            if (string.IsNullOrEmpty(editorPath))
            {
                // Fall back to default system editor
                try
                {
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(selected.FilePath)
                        { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening file:\n" + ex.Message,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return;
            }

            // Resolve relative path against app directory
            string resolvedPath = editorPath;
            if (!Path.IsPathRooted(editorPath))
            {
                resolvedPath = Path.Combine(appDir, editorPath);
            }

            if (!File.Exists(resolvedPath))
            {
                MessageBox.Show("Editor not found:\n" + resolvedPath +
                    "\n\nSet the editor path in App Settings, or leave blank to use the system default.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(
                    new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = resolvedPath,
                        Arguments = "\"" + selected.FilePath + "\"",
                        UseShellExecute = false
                    });
                statusFile.Text = "Opened in editor: " + Path.GetFileName(selected.FilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error launching editor:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseEditor_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Title = "Select Text Editor";
            dlg.Filter = "Executables (*.exe)|*.exe|All Files (*.*)|*.*";
            dlg.InitialDirectory = appDir;

            if (dlg.ShowDialog() == true)
            {
                string selectedPath = dlg.FileName;

                // Try to make it relative to app directory
                try
                {
                    string appDirFull = Path.GetFullPath(appDir).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
                    string selectedFull = Path.GetFullPath(selectedPath);

                    if (selectedFull.StartsWith(appDirFull, StringComparison.OrdinalIgnoreCase))
                    {
                        selectedPath = selectedFull.Substring(appDirFull.Length);
                    }
                }
                catch (Exception) { }

                editEditorPath.Text = selectedPath;
                windowSettings.EditorPath = selectedPath;
                statusFile.Text = "Editor set: " + selectedPath;
            }
        }

        private void SaveTrack_Click(object sender, RoutedEventArgs e)
        {
            // Use stored reference instead of relying on list selection
            TrackItem selected = currentTrackItem;

            if (selected == null)
            {
                CreateNewTrack();
                return;
            }

            string newName = editName.Text.Trim();

            if (string.IsNullOrWhiteSpace(newName))
            {
                MessageBox.Show("Program Name cannot be empty.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Check if name changed - need to rename the file
            string oldName = Path.GetFileNameWithoutExtension(selected.FilePath);

            // Update all fields from UI
            selected.ProgramName = newName;
            selected.TrackURL = editTrackURL.Text;
            selected.StartString = editStartString.Text;
            selected.StopString = editStopString.Text;
            selected.DownloadURL = editDownloadURL.Text;
            selected.Version = editVersion.Text;
            selected.ReleaseDate = editReleaseDate.Text;
            selected.PublisherName = editPublisherName.Text;
            selected.SuiteName = editSuiteName.Text;

            try
            {
                // If the name changed, rename the file
                if (oldName != newName)
                {
                    string safeName = newName;
                    foreach (char c in Path.GetInvalidFileNameChars())
                    {
                        safeName = safeName.Replace(c.ToString(), "");
                    }

                    string dir = Path.GetDirectoryName(selected.FilePath);
                    string newPath = Path.Combine(dir, safeName + ".track");

                    if (File.Exists(newPath) && newPath != selected.FilePath)
                    {
                        MessageBox.Show("A track file with that name already exists.",
                            "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }

                    if (File.Exists(selected.FilePath))
                    {
                        File.Delete(selected.FilePath);
                    }

                    selected.FilePath = newPath;
                }

                // If status was "changed" and version is now up to date, mark as unchanged
                if (selected.TrackStatus == "changed" &&
                    !string.IsNullOrEmpty(selected.LatestVersion) &&
                    selected.Version == selected.LatestVersion)
                {
                    selected.TrackStatus = "unchanged";
                }

                selected.SaveToFile();

                int selectedIndex = currentItems.IndexOf(selected);
                suppressAutoDownload = true;
                itemList.ItemsSource = null;
                itemList.ItemsSource = currentItems;
                if (selectedIndex >= 0)
                    itemList.SelectedIndex = selectedIndex;
                suppressAutoDownload = false;

                UpdateVersionDisplay();
                RefreshCategoryTree();
                // Re-download source for the current track
                if (currentTrackItem != null && !string.IsNullOrWhiteSpace(currentTrackItem.TrackURL))
                {
                    AutoDownloadSource(currentTrackItem.TrackURL);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving track: " + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DeleteTrack_Click(object sender, RoutedEventArgs e)
        {
            TrackItem selected = itemList.SelectedItem as TrackItem;

            if (selected == null)
            {
                MessageBox.Show("No track selected.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            MessageBoxResult result = MessageBox.Show(
                "Delete track '" + selected.ProgramName + "'?\n\n" +
                "File: " + Path.GetFileName(selected.FilePath),
                "Delete Track",
                MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    if (File.Exists(selected.FilePath))
                    {
                        File.Delete(selected.FilePath);
                    }

                    // Reload the current category
                    TreeViewItem selectedCat = categoryTree.SelectedItem as TreeViewItem;
                    if (selectedCat != null)
                    {
                        LoadTrackFiles((string)selectedCat.Tag);
                    }

                    RefreshCategoryTree();
                    ClearTrackFields();
                    statusFile.Text = "Deleted: " + selected.ProgramName;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error deleting track: " + ex.Message,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void CreateNewTrack()
        {
            // Use stored category path instead of relying on tree selection
            TreeViewItem selectedCat = categoryTree.SelectedItem as TreeViewItem;
            string folderPath = selectedCat != null
                ? (string)selectedCat.Tag
                : currentCategoryPath;

            if (string.IsNullOrEmpty(folderPath))
            {
                MessageBox.Show("Please select a category first.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string name = editName.Text;

            if (string.IsNullOrWhiteSpace(name))
            {
                name = ShowInputDialog("New Track", "Enter program name:");
                if (string.IsNullOrWhiteSpace(name))
                    return;
            }

            string safeName = name;
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(c.ToString(), "");
            }

            string filePath = Path.Combine(folderPath, safeName + ".track");

            if (File.Exists(filePath))
            {
                MessageBoxResult overwrite = MessageBox.Show(
                    "Track file '" + safeName + ".track' already exists.\n\n" +
                    "Do you want to overwrite it?",
                    "PAT v7", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (overwrite != MessageBoxResult.Yes)
                    return;
            }

            TrackItem newItem = new TrackItem();
            newItem.ProgramName = name;
            newItem.FilePath = filePath;
            newItem.TrackURL = editTrackURL.Text;
            newItem.StartString = editStartString.Text;
            newItem.StopString = editStopString.Text;
            newItem.DownloadURL = editDownloadURL.Text;
            newItem.Version = editVersion.Text;
            newItem.ReleaseDate = editReleaseDate.Text;
            newItem.PublisherName = editPublisherName.Text;
            newItem.SuiteName = editSuiteName.Text;
            newItem.CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

            try
            {
                newItem.SaveToFile(filePath);
                LoadTrackFiles(folderPath);
                RefreshCategoryTree();
                statusFile.Text = "Created: " + name;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating track: " + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        private void ClearTrackFields()
        {
            currentTrackItem = null;
            editName.Text = "";
            editTrackURL.Text = "";
            editStartString.Text = "";
            editStopString.Text = "";
            editDownloadURL.Text = "";
            editVersion.Text = "";
            editLatestVersion.Text = "";
            editLatestVersion.Foreground = Brushes.Gray;
            btnUpdateVersion.IsEnabled = false;
            editReleaseDate.Text = "";
            editPublisherName.Text = "";
            editSuiteName.Text = "";
        }

        private void UpdateVersion_Click(object sender, RoutedEventArgs e)
        {
            if (currentTrackItem != null && !string.IsNullOrEmpty(currentTrackItem.LatestVersion))
            {
                editVersion.Text = currentTrackItem.LatestVersion;
                currentTrackItem.Version = currentTrackItem.LatestVersion;
                currentTrackItem.SaveToFile();

                // Refresh list
                int selectedIndex = currentItems.IndexOf(currentTrackItem);
                suppressAutoDownload = true;
                itemList.ItemsSource = null;
                itemList.ItemsSource = currentItems;
                if (selectedIndex >= 0)
                    itemList.SelectedIndex = selectedIndex;
                suppressAutoDownload = false;

                UpdateVersionDisplay();
                statusFile.Text = "Version updated to " + currentTrackItem.Version;
            }
        }

        private void UpdateVersionDisplay()
        {
            string current = editVersion.Text.Trim();
            string latest = currentTrackItem?.LatestVersion ?? "";

            if (string.IsNullOrEmpty(latest) || latest == "(not yet checked)")
            {
                editLatestVersion.Text = "(not yet checked)";
                editLatestVersion.Foreground = Brushes.Gray;
                btnUpdateVersion.IsEnabled = false;
                return;
            }

            editLatestVersion.Text = latest;

            if (current == latest)
            {
                // Same version — green, button disabled
                editLatestVersion.Foreground = new SolidColorBrush(Color.FromRgb(0, 160, 0));
                btnUpdateVersion.IsEnabled = false;
            }
            else
            {
                // Different version — red, button enabled
                editLatestVersion.Foreground = new SolidColorBrush(Color.FromRgb(200, 30, 30));
                btnUpdateVersion.IsEnabled = true;
            }
        }

        // ============================
        // Layout Toggle
        // ============================

        private void ToggleLayout_Click(object sender, RoutedEventArgs e)
        {
            // Save current layout's splitter positions before switching
            SaveCurrentSplitterPositions();

            isHorizontalLayout = !isHorizontalLayout;

			if (isHorizontalLayout)
			{
			    layoutVertical.Visibility = Visibility.Collapsed;
			    layoutHorizontal.Visibility = Visibility.Visible;
			    btnToggleLayout.Content = "\xeca5";
			    btnToggleLayout.FontFamily = new FontFamily("Segoe Fluent Icons");
			    btnToggleLayout.FontSize = 16;
			    btnToggleLayout.ToolTip = "Switch to Vertical Layout";
			}
			else
			{
			    layoutHorizontal.Visibility = Visibility.Collapsed;
			    layoutVertical.Visibility = Visibility.Visible;
			    btnToggleLayout.Content = "\xeca5";
			    btnToggleLayout.FontFamily = new FontFamily("Segoe Fluent Icons");
			    btnToggleLayout.FontSize = 16;
			    btnToggleLayout.ToolTip = "Switch to Horizontal Layout";
			}

            PlaceControlsInLayout();

            // Restore the target layout's splitter positions after a brief delay
            this.Dispatcher.BeginInvoke(new Action(() =>
            {
                ApplySplitterPositions();
            }), System.Windows.Threading.DispatcherPriority.Loaded);

            statusFile.Text = isHorizontalLayout ? "Layout: Horizontal" : "Layout: Vertical";
        }

        private void SaveCurrentSplitterPositions()
        {
            try
            {
                if (!isHorizontalLayout && layoutVertical.IsVisible)
                {
                    var cols = layoutVertical.ColumnDefinitions;
                    if (cols.Count >= 3 && cols[0].ActualWidth > 0)
                        windowSettings.VSplitter1 = cols[0].ActualWidth;

                    foreach (var child in layoutVertical.Children)
                    {
                        if (child is Grid leftPanel && Grid.GetColumn((System.Windows.UIElement)child) == 0)
                        {
                            if (leftPanel.RowDefinitions.Count >= 3 && leftPanel.RowDefinitions[0].ActualHeight > 0)
                                windowSettings.VSplitter1Row = leftPanel.RowDefinitions[0].ActualHeight;

                            foreach (var sub in leftPanel.Children)
                            {
                                if (sub is Grid topGrid && Grid.GetRow((System.Windows.UIElement)sub) == 0)
                                {
                                    if (topGrid.ColumnDefinitions.Count >= 3 && topGrid.ColumnDefinitions[0].ActualWidth > 0)
                                        windowSettings.VSplitterCat = topGrid.ColumnDefinitions[0].ActualWidth;
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
                else if (isHorizontalLayout && layoutHorizontal.IsVisible)
                {
                    var rows = layoutHorizontal.RowDefinitions;
                    if (rows.Count >= 3 && rows[0].ActualHeight > 0)
                        windowSettings.HRowTop = rows[0].ActualHeight;

                    foreach (var child in layoutHorizontal.Children)
                    {
                        if (child is Grid g)
                        {
                            int row = Grid.GetRow((System.Windows.UIElement)child);
                            if (row == 0 && g.ColumnDefinitions.Count >= 3 && g.ColumnDefinitions[0].ActualWidth > 0)
                                windowSettings.HTopCol0 = g.ColumnDefinitions[0].ActualWidth;
                            else if (row == 2 && g.ColumnDefinitions.Count >= 3 && g.ColumnDefinitions[0].ActualWidth > 0)
                                windowSettings.HBotCol0 = g.ColumnDefinitions[0].ActualWidth;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
        }

        private void ApplySplitterPositions()
        {
            try
            {
                if (!isHorizontalLayout && layoutVertical.IsVisible)
                {
                    var cols = layoutVertical.ColumnDefinitions;

                    // Main left/right split (column 0)
                    if (!double.IsNaN(windowSettings.VSplitter1) && windowSettings.VSplitter1 > 0)
                        cols[0].Width = new GridLength(windowSettings.VSplitter1);

                    // Inside the left panel: top row has Cat+List
                    Grid leftPanel = layoutVertical.Children[0] as Grid;
                    if (leftPanel == null)
                    {
                        // Find the Grid in column 0
                        foreach (var child in layoutVertical.Children)
                        {
                            if (child is Grid g && Grid.GetColumn((System.Windows.UIElement)child) == 0)
                            {
                                leftPanel = g;
                                break;
                            }
                        }
                    }

                    if (leftPanel != null)
                    {
                        if (!double.IsNaN(windowSettings.VSplitter1Row) && windowSettings.VSplitter1Row >= 150)
                        {
                            // Use star ratios instead of fixed pixel heights
                            // This lets the grid enforce MinHeight properly
                            double totalHeight = leftPanel.ActualHeight;
                            if (totalHeight <= 0) totalHeight = this.ActualHeight - 100;
                            double trackSettingsHeight = totalHeight - windowSettings.VSplitter1Row - 5;

                            // Ensure minimum height for track settings
                            if (trackSettingsHeight < 250) trackSettingsHeight = 250;
                            double topHeight = totalHeight - trackSettingsHeight - 5;
                            if (topHeight < 150) topHeight = 150;

                            // Use star sizing so grid can enforce mins on resize
                            leftPanel.RowDefinitions[0].Height = new GridLength(topHeight, GridUnitType.Star);
                            leftPanel.RowDefinitions[2].Height = new GridLength(trackSettingsHeight, GridUnitType.Star);
                        }

                        // Cat/List splitter
                        foreach (var child in leftPanel.Children)
                        {
                            if (child is Grid topGrid && Grid.GetRow((System.Windows.UIElement)child) == 0)
                            {
                                if (topGrid.ColumnDefinitions.Count >= 3 &&
                                    !double.IsNaN(windowSettings.VSplitterCat) &&
                                    windowSettings.VSplitterCat > 0)
                                {
                                    topGrid.ColumnDefinitions[0].Width = new GridLength(windowSettings.VSplitterCat);
                                }
                                break;
                            }
                        }
                    }
                }
                else if (isHorizontalLayout && layoutHorizontal.IsVisible)
                {
                    // Horizontal layout
                    var rows = layoutHorizontal.RowDefinitions;

                    if (!double.IsNaN(windowSettings.HRowTop) && windowSettings.HRowTop > 0)
                    {
                        double totalHeight = this.ActualHeight - 100;
                        if (totalHeight <= 0) totalHeight = 600;
                        double topHeight = windowSettings.HRowTop;
                        double bottomHeight = totalHeight - topHeight - 5;

                        if (bottomHeight < 250) bottomHeight = 250;
                        topHeight = totalHeight - bottomHeight - 5;
                        if (topHeight < 150) topHeight = 150;

                        // Use star sizing so grid enforces MinHeight on resize
                        rows[0].Height = new GridLength(topHeight, GridUnitType.Star);
                        rows[2].Height = new GridLength(bottomHeight, GridUnitType.Star);
                    }

                    // Top row columns
                    Grid topRow = null;
                    foreach (var child in layoutHorizontal.Children)
                    {
                        if (child is Grid g && Grid.GetRow((System.Windows.UIElement)child) == 0)
                        {
                            topRow = g;
                            break;
                        }
                    }

                    if (topRow != null && topRow.ColumnDefinitions.Count >= 3 &&
                        !double.IsNaN(windowSettings.HTopCol0) && windowSettings.HTopCol0 > 0)
                    {
                        topRow.ColumnDefinitions[0].Width = new GridLength(windowSettings.HTopCol0);
                    }

                    // Bottom row columns
                    Grid botRow = null;
                    foreach (var child in layoutHorizontal.Children)
                    {
                        if (child is Grid g && Grid.GetRow((System.Windows.UIElement)child) == 2)
                        {
                            botRow = g;
                            break;
                        }
                    }

                    if (botRow != null && botRow.ColumnDefinitions.Count >= 3 &&
                        !double.IsNaN(windowSettings.HBotCol0) && windowSettings.HBotCol0 > 0)
                    {
                        botRow.ColumnDefinitions[0].Width = new GridLength(windowSettings.HBotCol0);
                    }
                }
            }
            catch (Exception)
            {
                // Silently ignore layout restore errors
            }
        }

        private void SaveWindowSettings()
        {
            try
            {
                if (this.WindowState == WindowState.Maximized)
                {
                    windowSettings.IsMaximized = true;
                }
                else
                {
                    windowSettings.IsMaximized = false;
                    windowSettings.WindowLeft = this.Left;
                    windowSettings.WindowTop = this.Top;
                    windowSettings.WindowWidth = this.Width;
                    windowSettings.WindowHeight = this.Height;
                }

                windowSettings.IsHorizontalLayout = isHorizontalLayout;

                SaveCurrentSplitterPositions();
                windowSettings.SourceFontSize = sourceView.FontSize;
                SaveColumnSettings();
                windowSettings.EditorPath = editEditorPath.Text.Trim();
                windowSettings.Save(windowSettingsPath);
            }
            catch (Exception)
            {
            }
        }

		// ============================
		// Theme Application
		// ============================

		private TabItem BuildThemeTab()
		{
		    TabItem themeTab = new TabItem();
		    themeTab.Header = "🎨 Theme";

		    DockPanel themeOuterPanel = new DockPanel();

		    // ---- Declare themePanel EARLY so lambda closures can reference it ----
		    StackPanel themePanel = new StackPanel();
		    themePanel.Margin = new Thickness(12);

		    // ---- Toolbar row ----
		    ToolBarTray themeToolBarTray = new ToolBarTray();
		    DockPanel.SetDock(themeToolBarTray, Dock.Top);

		    ToolBar themeToolBar = new ToolBar();
		    themeToolBar.Band = 0;
		    themeToolBar.BandIndex = 0;

		    TextBlock themeToolBarTitle = new TextBlock();
		    themeToolBarTitle.Text = "Theme Settings";
		    themeToolBarTitle.FontWeight = FontWeights.Bold;
		    themeToolBarTitle.FontSize = 13;
		    themeToolBarTitle.VerticalAlignment = VerticalAlignment.Center;
		    themeToolBarTitle.Margin = new Thickness(0, 0, 8, 0);
		    themeToolBar.Items.Add(themeToolBarTitle);

		    themeToolBar.Items.Add(new Separator());

		    ComboBox presetCombo = new ComboBox();
		    presetCombo.Width = 180;
		    // Add built-in presets
		    presetCombo.Items.Add("Default Light");
		    presetCombo.Items.Add("Dark");
		    presetCombo.Items.Add("Blue");
		    // Add custom .theme files from Settings folder
		    PopulateThemeCombo(presetCombo);
		    presetCombo.SelectedIndex = 0;

		    // Instant-apply on selection change (no separate "Apply" button needed)
			// Hook DropDownOpened once to style separator items
			presetCombo.DropDownOpened += (ds, de) =>
			{
			    for (int i = 0; i < presetCombo.Items.Count; i++)
			    {
			        if (presetCombo.Items[i] is Separator)
			        {
			            ComboBoxItem container = presetCombo.ItemContainerGenerator
			                .ContainerFromIndex(i) as ComboBoxItem;
			            if (container != null)
			            {
			                container.IsEnabled = false;
			                container.Padding = new Thickness(0);
			                container.MinHeight = 0;
			                container.Height = 2;
			            }
			        }
			    }
			};

			presetCombo.SelectionChanged += (s, ev) =>
			{
			    // If user clicked a Separator, revert to previous selection
			    if (presetCombo.SelectedItem is Separator || presetCombo.SelectedItem == null)
			    {
			        if (ev.RemovedItems.Count > 0 && ev.RemovedItems[0] is string prev)
			            presetCombo.SelectedItem = prev;
			        else
			            presetCombo.SelectedIndex = 0;
			        return;
			    }

			    string selected = presetCombo.SelectedItem as string;
			    if (string.IsNullOrEmpty(selected)) return;

			    switch (selected)
			    {
			        case "Default Light":
			            previewTheme = ThemeSettings.GetDefaultLight();
			            break;
			        case "Dark":
			            previewTheme = ThemeSettings.GetDarkTheme();
			            break;
			        case "Blue":
			            previewTheme = ThemeSettings.GetBlueTheme();
			            break;
			        default:
			            string themePath = Path.Combine(settingsPath, selected + ".theme");
			            if (File.Exists(themePath))
			            {
			                previewTheme = ThemeSettings.Load(themePath);
			                currentThemePath = themePath;
			                windowSettings.ActiveThemePath = selected + ".theme";
			            }
			            else
			            {
			                previewTheme = ThemeSettings.GetDefaultLight();
			            }
			            break;
			    }
			    ApplyTheme(previewTheme);
			    RefreshThemeSwatches(themePanel);
			    statusFile.Text = "Theme applied: " + selected;
			};
			
		    themeToolBar.Items.Add(presetCombo);

		    Button btnApplyPreset = new Button();
		    btnApplyPreset.Content = "\xe930";
		    btnApplyPreset.FontFamily = new FontFamily("Segoe Fluent Icons");
		    btnApplyPreset.FontSize = 16;
		    btnApplyPreset.Padding = new Thickness(6, 2, 6, 2);
		    btnApplyPreset.ToolTip = "Apply preset theme";
		    themeToolBar.Items.Add(btnApplyPreset);

		    themeToolBar.Items.Add(new Separator());

		    Button btnLoadTheme = new Button();
		    btnLoadTheme.Content = "\xe838";
		    btnLoadTheme.FontFamily = new FontFamily("Segoe Fluent Icons");
		    btnLoadTheme.FontSize = 16;
		    btnLoadTheme.Padding = new Thickness(6, 2, 6, 2);
		    btnLoadTheme.ToolTip = "Browse for theme file";
		    themeToolBar.Items.Add(btnLoadTheme);

		    Button btnSaveTheme = new Button();
		    btnSaveTheme.Content = "\uE74E";
		    btnSaveTheme.FontFamily = new FontFamily("Segoe Fluent Icons");
		    btnSaveTheme.FontSize = 16;
		    btnSaveTheme.Padding = new Thickness(6, 2, 6, 2);
		    btnSaveTheme.ToolTip = "Save theme to file";
		    themeToolBar.Items.Add(btnSaveTheme);

			Button btnSaveThemeAs = new Button();
			btnSaveThemeAs.Content = "\uE792"; // SaveAs icon
			btnSaveThemeAs.FontFamily = new FontFamily("Segoe Fluent Icons");
			btnSaveThemeAs.FontSize = 16;
			btnSaveThemeAs.Padding = new Thickness(6, 2, 6, 2);
			btnSaveThemeAs.ToolTip = "Save theme as new file";
			themeToolBar.Items.Add(btnSaveThemeAs);
			
		    themeToolBar.Items.Add(new Separator());

		    Button btnResetTheme = new Button();
		    btnResetTheme.Content = "\xe845";
		    btnResetTheme.FontFamily = new FontFamily("Segoe Fluent Icons");
		    btnResetTheme.FontSize = 16;
		    btnResetTheme.Padding = new Thickness(6, 2, 6, 2);
		    btnResetTheme.ToolTip = "Revert to last saved/loaded theme";
		    themeToolBar.Items.Add(btnResetTheme);

		    themeToolBarTray.ToolBars.Add(themeToolBar);
		    themeOuterPanel.Children.Add(themeToolBarTray);

		    // ---- Scrollable color swatch area ----
		    ScrollViewer themeScroll = new ScrollViewer();
		    themeScroll.VerticalScrollBarVisibility = ScrollBarVisibility.Auto;

		    // themePanel was declared at the top of the method

		    // ---- Build color swatch rows ----
		    AddThemeSection(themePanel, "Window & General");
		    AddColorSwatch(themePanel, "Window Background", () => previewTheme.WindowBackground, c => previewTheme.WindowBackground = c);
		    AddColorSwatch(themePanel, "Window Foreground", () => previewTheme.WindowForeground, c => previewTheme.WindowForeground = c);
		    AddColorSwatch(themePanel, "Button Background", () => previewTheme.ButtonBackground, c => previewTheme.ButtonBackground = c);
		    AddColorSwatch(themePanel, "Button Foreground", () => previewTheme.ButtonForeground, c => previewTheme.ButtonForeground = c);
		    AddColorSwatch(themePanel, "Splitter Color", () => previewTheme.SplitterColor, c => previewTheme.SplitterColor = c);

		    AddThemeSection(themePanel, "Menu & Toolbar");
		    AddColorSwatch(themePanel, "Menu Background", () => previewTheme.MenuBackground, c => previewTheme.MenuBackground = c);
		    AddColorSwatch(themePanel, "Menu Foreground", () => previewTheme.MenuForeground, c => previewTheme.MenuForeground = c);
		    AddColorSwatch(themePanel, "Toolbar Background", () => previewTheme.ToolbarBackground, c => previewTheme.ToolbarBackground = c);
		    AddColorSwatch(themePanel, "Toolbar Foreground", () => previewTheme.ToolbarForeground, c => previewTheme.ToolbarForeground = c);

		    AddThemeSection(themePanel, "Status Bar");
		    AddColorSwatch(themePanel, "StatusBar Background", () => previewTheme.StatusBarBackground, c => previewTheme.StatusBarBackground = c);
		    AddColorSwatch(themePanel, "StatusBar Foreground", () => previewTheme.StatusBarForeground, c => previewTheme.StatusBarForeground = c);

		    AddThemeSection(themePanel, "Category Tree");
		    AddColorSwatch(themePanel, "Tree Background", () => previewTheme.TreeBackground, c => previewTheme.TreeBackground = c);
		    AddColorSwatch(themePanel, "Tree Foreground", () => previewTheme.TreeForeground, c => previewTheme.TreeForeground = c);
		    AddColorSwatch(themePanel, "Tree Selected Background", () => previewTheme.TreeSelectedBackground, c => previewTheme.TreeSelectedBackground = c);
		    AddColorSwatch(themePanel, "Tree Selected Foreground", () => previewTheme.TreeSelectedForeground, c => previewTheme.TreeSelectedForeground = c);

		    AddThemeSection(themePanel, "ListView");
		    AddColorSwatch(themePanel, "List Background", () => previewTheme.ListBackground, c => previewTheme.ListBackground = c);
		    AddColorSwatch(themePanel, "List: Name/Version Columns", () => previewTheme.ListForeground, c => previewTheme.ListForeground = c);
		    AddColorSwatch(themePanel, "List Header Background", () => previewTheme.ListHeaderBackground, c => previewTheme.ListHeaderBackground = c);
		    AddColorSwatch(themePanel, "List Header Foreground", () => previewTheme.ListHeaderForeground, c => previewTheme.ListHeaderForeground = c);
		    AddColorSwatch(themePanel, "List Selected Background", () => previewTheme.ListSelectedBackground, c => previewTheme.ListSelectedBackground = c);
		    AddColorSwatch(themePanel, "List Selected Foreground", () => previewTheme.ListSelectedForeground, c => previewTheme.ListSelectedForeground = c);
		    AddColorSwatch(themePanel, "List Hover Background", () => previewTheme.ListHoverBackground, c => previewTheme.ListHoverBackground = c);
		    AddColorSwatch(themePanel, "List Hover Foreground", () => previewTheme.ListHoverForeground, c => previewTheme.ListHoverForeground = c);

		    AddThemeSection(themePanel, "List Status Indicators (● icons)");
		    AddColorSwatch(themePanel, "Unchanged", () => previewTheme.StatusUnchanged, c => previewTheme.StatusUnchanged = c);
		    AddColorSwatch(themePanel, "Changed", () => previewTheme.StatusChanged, c => previewTheme.StatusChanged = c);
		    AddColorSwatch(themePanel, "Error", () => previewTheme.StatusError, c => previewTheme.StatusError = c);
		    AddColorSwatch(themePanel, "New", () => previewTheme.StatusNew, c => previewTheme.StatusNew = c);
		    AddColorSwatch(themePanel, "Unchecked", () => previewTheme.StatusUnchecked, c => previewTheme.StatusUnchecked = c);

		    AddThemeSection(themePanel, "List Cell Status Colors (text color for URL, Hash, etc.)");
		    AddColorSwatch(themePanel, "OK", () => previewTheme.CellStatusOk, c => previewTheme.CellStatusOk = c);
		    AddColorSwatch(themePanel, "Changed", () => previewTheme.CellStatusChanged, c => previewTheme.CellStatusChanged = c);
		    AddColorSwatch(themePanel, "Error", () => previewTheme.CellStatusError, c => previewTheme.CellStatusError = c);
		    AddColorSwatch(themePanel, "Default", () => previewTheme.CellStatusDefault, c => previewTheme.CellStatusDefault = c);

		    AddThemeSection(themePanel, "Track Settings Panel");
		    AddColorSwatch(themePanel, "Panel Background", () => previewTheme.PanelBackground, c => previewTheme.PanelBackground = c);
		    AddColorSwatch(themePanel, "TextBox Background", () => previewTheme.TextBoxBackground, c => previewTheme.TextBoxBackground = c);
		    AddColorSwatch(themePanel, "TextBox Foreground", () => previewTheme.TextBoxForeground, c => previewTheme.TextBoxForeground = c);

		    AddThemeSection(themePanel, "Tabs");
		    AddColorSwatch(themePanel, "Tab Background", () => previewTheme.TabBackground, c => previewTheme.TabBackground = c);
		    AddColorSwatch(themePanel, "Tab Foreground", () => previewTheme.TabForeground, c => previewTheme.TabForeground = c);
		    AddColorSwatch(themePanel, "Tab Selected Background", () => previewTheme.TabSelectedBackground, c => previewTheme.TabSelectedBackground = c);
		    AddColorSwatch(themePanel, "Tab Selected Foreground", () => previewTheme.TabSelectedForeground, c => previewTheme.TabSelectedForeground = c);

		    AddThemeSection(themePanel, "Source View");
		    AddColorSwatch(themePanel, "Source Background", () => previewTheme.SourceBackground, c => previewTheme.SourceBackground = c);
		    AddColorSwatch(themePanel, "Start String Color", () => previewTheme.SourceStartStringColor, c => previewTheme.SourceStartStringColor = c);
		    AddColorSwatch(themePanel, "Info String Color", () => previewTheme.SourceInfoStringColor, c => previewTheme.SourceInfoStringColor = c);
		    AddColorSwatch(themePanel, "Stop String Color", () => previewTheme.SourceStopStringColor, c => previewTheme.SourceStopStringColor = c);
		    AddColorSwatch(themePanel, "HTML Tag Color", () => previewTheme.SourceTagColor, c => previewTheme.SourceTagColor = c);
		    AddColorSwatch(themePanel, "Text Content Color", () => previewTheme.SourceTextColor, c => previewTheme.SourceTextColor = c);

		    AddThemeSection(themePanel, "Scrollbars");
		    AddColorSwatch(themePanel, "Scrollbar Background", () => previewTheme.ScrollBarBackground, c => previewTheme.ScrollBarBackground = c);
		    AddColorSwatch(themePanel, "Scrollbar Thumb", () => previewTheme.ScrollBarThumb, c => previewTheme.ScrollBarThumb = c);

		    // ---- Button handlers ----

			btnApplyPreset.Click += (s, ev) =>
			{
			    // Skip if a Separator is somehow selected
			    if (presetCombo.SelectedItem is Separator || presetCombo.SelectedItem == null)
			        return;

			    string selected = presetCombo.SelectedItem as string;
			    if (string.IsNullOrEmpty(selected))
			        return;

			    switch (selected)
			    {
			        case "Default Light":
			            previewTheme = ThemeSettings.GetDefaultLight();
			            break;
			        case "Dark":
			            previewTheme = ThemeSettings.GetDarkTheme();
			            break;
			        case "Blue":
			            previewTheme = ThemeSettings.GetBlueTheme();
			            break;
			        default:
			            // User theme from Settings folder
			            string themePath = Path.Combine(settingsPath, selected + ".theme");
			            if (File.Exists(themePath))
			            {
			                previewTheme = ThemeSettings.Load(themePath);
			                currentThemePath = themePath;
			                windowSettings.ActiveThemePath = selected + ".theme";
			            }
			            else
			            {
			                previewTheme = ThemeSettings.GetDefaultLight();
			            }
			            break;
			    }
			    ApplyTheme(previewTheme);
			    RefreshThemeSwatches(themePanel);
			    statusFile.Text = "Theme applied: " + selected;
			};

		    btnLoadTheme.Click += (s, ev) =>
		    {
		        Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
		        dlg.Title = "Load Theme";
		        dlg.Filter = "Theme Files (*.theme)|*.theme|XML Files (*.xml)|*.xml|All Files (*.*)|*.*";
		        dlg.InitialDirectory = settingsPath;

		        if (dlg.ShowDialog() == true)
		        {
		            ThemeSettings loaded = ThemeSettings.Load(dlg.FileName);
		            previewTheme = loaded;
		            currentThemePath = dlg.FileName;
		            ApplyTheme(previewTheme);
		            RefreshThemeSwatches(themePanel);

		            string relativePath = Path.GetFileName(dlg.FileName);
		            windowSettings.ActiveThemePath = relativePath;

		            // Refresh the combo to include the newly loaded file if it's in Settings
		            PopulateThemeCombo(presetCombo);

		            statusFile.Text = "Theme loaded: " + loaded.ThemeName;
		        }
		    };

			btnSaveTheme.Click += (s, ev) =>
			{
			    // If we have a current theme file path, save silently
			    if (!string.IsNullOrEmpty(currentThemePath) && File.Exists(currentThemePath))
			    {
			        previewTheme.Save(currentThemePath);
			        currentTheme = previewTheme.Clone();
			        statusFile.Text = "Theme saved: " + Path.GetFileNameWithoutExtension(currentThemePath);
			        return;
			    }

			    // No current file — check if a custom theme is selected in the combo
			    string selected = presetCombo.SelectedItem as string;
			    if (!string.IsNullOrEmpty(selected) &&
			        selected != "Default Light" && selected != "Dark" && selected != "Blue")
			    {
			        string savePath = Path.Combine(settingsPath, selected + ".theme");
			        previewTheme.ThemeName = selected;
			        previewTheme.Save(savePath);
			        currentThemePath = savePath;
			        currentTheme = previewTheme.Clone();
			        windowSettings.ActiveThemePath = selected + ".theme";
			        statusFile.Text = "Theme saved: " + selected;
			        return;
			    }

			    // Built-in theme selected or no file — fall back to Save As behavior
			    // (Can't overwrite built-in presets, so prompt for a name)
			    string themeName = ShowInputDialog("Save Theme As", "Enter a name for this theme:");
			    if (string.IsNullOrWhiteSpace(themeName)) return;

			    string safeName = themeName;
			    foreach (char c in Path.GetInvalidFileNameChars())
			        safeName = safeName.Replace(c.ToString(), "");

			    if (string.IsNullOrWhiteSpace(safeName))
			    {
			        MessageBox.Show("Invalid theme name.", "PAT v7",
			            MessageBoxButton.OK, MessageBoxImage.Warning);
			        return;
			    }

			    previewTheme.ThemeName = themeName;
			    string newPath = Path.Combine(settingsPath, safeName + ".theme");
			    previewTheme.Save(newPath);
			    currentThemePath = newPath;
			    currentTheme = previewTheme.Clone();
			    windowSettings.ActiveThemePath = safeName + ".theme";
			    PopulateThemeCombo(presetCombo);
			    statusFile.Text = "Theme saved: " + themeName;
			};

			btnSaveThemeAs.Click += (s, ev) =>
			{
			    string themeName = ShowInputDialog("Save Theme As", "Enter a name for this theme:");
			    if (string.IsNullOrWhiteSpace(themeName))
			        return;

			    string safeName = themeName;
			    foreach (char c in Path.GetInvalidFileNameChars())
			        safeName = safeName.Replace(c.ToString(), "");

			    if (string.IsNullOrWhiteSpace(safeName))
			    {
			        MessageBox.Show("Invalid theme name.", "PAT v7",
			            MessageBoxButton.OK, MessageBoxImage.Warning);
			        return;
			    }

			    previewTheme.ThemeName = themeName;
			    string savePath = Path.Combine(settingsPath, safeName + ".theme");

			    if (File.Exists(savePath))
			    {
			        MessageBoxResult overwrite = MessageBox.Show(
			            "Theme file '" + safeName + ".theme' already exists.\nOverwrite?",
			            "PAT v7", MessageBoxButton.YesNo, MessageBoxImage.Question);
			        if (overwrite != MessageBoxResult.Yes)
			            return;
			    }

			    previewTheme.Save(savePath);
			    currentThemePath = savePath;
			    currentTheme = previewTheme.Clone();
			    windowSettings.ActiveThemePath = safeName + ".theme";
			    PopulateThemeCombo(presetCombo);
			    statusFile.Text = "Theme saved as: " + themeName;
			};

		    btnResetTheme.Click += (s, ev) =>
		    {
		        if (!string.IsNullOrEmpty(currentThemePath) && File.Exists(currentThemePath))
		        {
		            previewTheme = ThemeSettings.Load(currentThemePath);
		        }
		        else
		        {
		            previewTheme = currentTheme.Clone();
		        }
		        ApplyTheme(previewTheme);
		        RefreshThemeSwatches(themePanel);
		        statusFile.Text = "Theme reset to saved state";
		    };

		    themeScroll.Content = themePanel;
		    themeOuterPanel.Children.Add(themeScroll);
		    themeTab.Content = themeOuterPanel;

		    // Initialize preview theme
		    previewTheme = currentTheme.Clone();

		    return themeTab;
		}

		private void PopulateThemeCombo(ComboBox combo)
		{
		    string currentSelection = combo.SelectedItem as string;

		    combo.Items.Clear();
		    combo.Items.Add("Default Light");
		    combo.Items.Add("Dark");
		    combo.Items.Add("Blue");

		    // Add custom .theme files from the Settings folder
		    if (Directory.Exists(settingsPath))
		    {
		        string[] themeFiles = Directory.GetFiles(settingsPath, "*.theme");
		        bool separatorAdded = false;

		        foreach (string file in themeFiles)
		        {
		            string name = Path.GetFileNameWithoutExtension(file);
		            if (name != "Default Light" && name != "Dark" && name != "Blue")
		            {
		                // Add separator before the first custom theme
		                if (!separatorAdded)
		                {
		                    combo.Items.Add(new Separator());
		                    separatorAdded = true;
		                }
		                combo.Items.Add(name);
		            }
		        }
		    }

		    if (!string.IsNullOrEmpty(currentSelection) && combo.Items.Contains(currentSelection))
		    {
		        combo.SelectedItem = currentSelection;
		    }
		    else
		    {
		        combo.SelectedIndex = 0;
		    }
		}

		private void ThemeAllComboBoxArrows(DependencyObject parent, SolidColorBrush arrowBrush)
		{
		    int count = VisualTreeHelper.GetChildrenCount(parent);
		    for (int i = 0; i < count; i++)
		    {
		        DependencyObject child = VisualTreeHelper.GetChild(parent, i);

		        if (child is System.Windows.Shapes.Path path &&
		            path.Name == "Arrow")
		        {
		            path.Fill = arrowBrush;
		        }

		        ThemeAllComboBoxArrows(child, arrowBrush);
		    }
		}
		
		private void AddThemeSection(StackPanel parent, string title)
		{
		    TextBlock header = new TextBlock();
		    header.Text = title;
		    header.FontWeight = FontWeights.Bold;
		    header.FontSize = 12;
		    header.Margin = new Thickness(0, 12, 0, 4);
		    parent.Children.Add(header);

		    Separator sep = new Separator();
		    sep.Margin = new Thickness(0, 0, 0, 4);
		    parent.Children.Add(sep);
		}

		private void AddColorSwatch(StackPanel parent, string label,
		    Func<Color> getter, Action<Color> setter)
		{
		    DockPanel row = new DockPanel();
		    row.Margin = new Thickness(0, 2, 0, 2);

		    // Clickable color swatch (acts as the pick button)
		    Border swatch = new Border();
		    swatch.Width = 28;
		    swatch.Height = 20;
		    swatch.BorderBrush = Brushes.Gray;
		    swatch.BorderThickness = new Thickness(1);
		    swatch.CornerRadius = new CornerRadius(3);
		    swatch.Background = new SolidColorBrush(getter());
		    swatch.Cursor = System.Windows.Input.Cursors.Hand;
		    swatch.ToolTip = "Click to pick a color";
		    swatch.Margin = new Thickness(0, 0, 8, 0);
		    DockPanel.SetDock(swatch, Dock.Left);

		    // Hex text display
		    TextBlock hexText = new TextBlock();
		    hexText.Text = ThemeSettings.ColorToHex(getter());
		    hexText.VerticalAlignment = VerticalAlignment.Center;
		    hexText.FontFamily = new FontFamily("Consolas");
		    hexText.FontSize = 11;
		    hexText.Foreground = Brushes.Gray;
		    hexText.Width = 60;
		    hexText.Margin = new Thickness(0, 0, 8, 0);
		    DockPanel.SetDock(hexText, Dock.Right);

		    // Label
		    TextBlock lbl = new TextBlock();
		    lbl.Text = label;
		    lbl.VerticalAlignment = VerticalAlignment.Center;

		    // Click handler — opens Windows color dialog
		    swatch.MouseLeftButtonDown += (s, ev) =>
		    {
		        Color currentColor = getter();

		        System.Windows.Forms.ColorDialog colorDialog = new System.Windows.Forms.ColorDialog();
		        colorDialog.Color = System.Drawing.Color.FromArgb(currentColor.R, currentColor.G, currentColor.B);
		        colorDialog.FullOpen = true;
		        colorDialog.AnyColor = true;

		        if (colorDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
		        {
		            Color newColor = Color.FromRgb(
		                colorDialog.Color.R,
		                colorDialog.Color.G,
		                colorDialog.Color.B);

		            setter(newColor);
		            swatch.Background = new SolidColorBrush(newColor);
		            hexText.Text = ThemeSettings.ColorToHex(newColor);

		            // Live preview
		            ApplyTheme(previewTheme);
		        }
		    };

		    row.Children.Add(swatch);
		    row.Children.Add(hexText);
		    row.Children.Add(lbl);
		    parent.Children.Add(row);

		    // Tag the swatch with getter/setter for refresh
		    swatch.Tag = new Tuple<Func<Color>, Action<Color>>(getter, setter);
		}

		private void RefreshThemeSwatches(StackPanel themePanel)
		{
		    RefreshSwatchesRecursive(themePanel);
		}

		private void RefreshSwatchesRecursive(object element)
		{
		    if (element is DockPanel dp)
		    {
		        foreach (var child in dp.Children)
		        {
		            if (child is Border swatch && swatch.Tag is Tuple<Func<Color>, Action<Color>> tag)
		            {
		                Color c = tag.Item1();
		                swatch.Background = new SolidColorBrush(c);

		                // Find the hex text in the same DockPanel
		                foreach (var sibling in dp.Children)
		                {
		                    if (sibling is TextBlock tb && tb.FontFamily != null &&
		                        tb.FontFamily.Source == "Consolas")
		                    {
		                        tb.Text = ThemeSettings.ColorToHex(c);
		                        break;
		                    }
		                }
		            }
		        }
		    }
		    else if (element is StackPanel sp)
		    {
		        foreach (var child in sp.Children)
		            RefreshSwatchesRecursive(child);
		    }
		}

		// ============================
		// Theme Application
		// ============================

		private void ApplyTheme(ThemeSettings theme)
		{
			this.Resources[SystemColors.HighlightBrushKey] = new SolidColorBrush(theme.MenuBackground);
			this.Resources[SystemColors.ActiveBorderBrushKey] = new SolidColorBrush(theme.MenuBackground);
		    currentTheme = theme;
		    StatusToColorConverter.UpdateFromTheme(theme);

		    SolidColorBrush winBg = new SolidColorBrush(theme.WindowBackground);
		    SolidColorBrush winFg = new SolidColorBrush(theme.WindowForeground);

		    // Window background
		    this.Background = winBg;
		    mainContent.Background = winBg;

		    // Layout grids and border hosts
		    layoutVertical.Background = winBg;
		    layoutHorizontal.Background = winBg;
		    catHostV.Background = winBg;
		    catHostH.Background = winBg;
		    listHostV.Background = winBg;
		    listHostH.Background = winBg;
		    trackSettingsHostV.Background = winBg;
		    trackSettingsHostH.Background = winBg;
		    tabsHostV.Background = winBg;
		    tabsHostH.Background = winBg;

			// ---- Menu (bar + dropdowns) ----
			SolidColorBrush menuBg = new SolidColorBrush(theme.MenuBackground);
			SolidColorBrush menuFg = new SolidColorBrush(theme.MenuForeground);
			foreach (var child in ((DockPanel)this.Content).Children)
			{
			    if (child is Menu menu)
			    {
			        menu.Background = menuBg;
			        menu.Foreground = menuFg;

			        foreach (var item in menu.Items)
			        {
			            if (item is MenuItem mi)
			                ApplyThemeToMenuItem(mi, menuBg, menuFg, true);
			        }
			    }
			}
			
			// ---- Toolbar ----
			SolidColorBrush tbBg = new SolidColorBrush(theme.ToolbarBackground);
			SolidColorBrush tbFg = new SolidColorBrush(theme.ToolbarForeground);
			ApplyThemeToToolBarTray(toolBarTray, tbBg, tbFg);

		    // ---- Status Bar ----
		    SolidColorBrush sbBg = new SolidColorBrush(theme.StatusBarBackground);
		    SolidColorBrush sbFg = new SolidColorBrush(theme.StatusBarForeground);
		    foreach (var child in ((DockPanel)this.Content).Children)
		    {
		        if (child is StatusBar sb)
		        {
		            sb.Background = sbBg;
		            sb.Foreground = sbFg;
		            statusFile.Foreground = sbFg;
		            ApplyStatusBarLegendColors(sb, theme);
		        }
		    }

		    // ---- Splitters ----
		    SolidColorBrush splitterBrush = new SolidColorBrush(theme.SplitterColor);
		    ApplyThemeToElement<GridSplitter>(layoutVertical, gs => gs.Background = splitterBrush);
		    ApplyThemeToElement<GridSplitter>(layoutHorizontal, gs => gs.Background = splitterBrush);
		    // Also get nested grids
		    foreach (var child in layoutVertical.Children)
		    {
		        if (child is Grid g)
		        {
		            ApplyThemeToElement<GridSplitter>(g, gs => gs.Background = splitterBrush);
		            foreach (var sub in g.Children)
		            {
		                if (sub is Grid sg)
		                    ApplyThemeToElement<GridSplitter>(sg, gs => gs.Background = splitterBrush);
		            }
		        }
		    }
		    foreach (var child in layoutHorizontal.Children)
		    {
		        if (child is Grid g)
		        {
		            ApplyThemeToElement<GridSplitter>(g, gs => gs.Background = splitterBrush);
		        }
		    }

			// ---- Category Tree ----
			categoryTree.Background = new SolidColorBrush(theme.TreeBackground);
			categoryTree.Foreground = new SolidColorBrush(theme.TreeForeground);
			categoryPanel.Background = new SolidColorBrush(theme.WindowBackground);

			// Also theme the category buttons at the bottom
			foreach (var child in categoryPanel.Children)
			{
			    if (child is StackPanel sp)
			    {
			        foreach (var btn in sp.Children)
			        {
			            if (btn is Button b)
			            {
			                // Buttons will pick up the global button style
			            }
			        }
			    }
			}

			// Override system color resources so the default template uses our colors
			this.Resources[SystemColors.HighlightBrushKey] = new SolidColorBrush(theme.TreeSelectedBackground);
			this.Resources[SystemColors.HighlightTextBrushKey] = new SolidColorBrush(theme.TreeSelectedForeground);
			this.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = new SolidColorBrush(theme.TreeSelectedBackground);
			this.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = new SolidColorBrush(theme.TreeSelectedForeground);

			Style treeItemStyle = new Style(typeof(TreeViewItem));
			treeItemStyle.Setters.Add(new Setter(TreeViewItem.ForegroundProperty, new SolidColorBrush(theme.TreeForeground)));
			treeItemStyle.Setters.Add(new Setter(TreeViewItem.BackgroundProperty, Brushes.Transparent));

			// Custom template so selection colors actually work
			ControlTemplate treeItemTemplate = new ControlTemplate(typeof(TreeViewItem));

			// Outer StackPanel: header row + items host
			FrameworkElementFactory outerStack = new FrameworkElementFactory(typeof(StackPanel));

			// The header row: border that binds to the TreeViewItem's Background
			FrameworkElementFactory headerBorder = new FrameworkElementFactory(typeof(Border));
			headerBorder.SetBinding(Border.BackgroundProperty,
			    new System.Windows.Data.Binding("Background")
			    {
			        RelativeSource = new System.Windows.Data.RelativeSource(
			            System.Windows.Data.RelativeSourceMode.TemplatedParent)
			    });
			headerBorder.SetValue(Border.PaddingProperty, new Thickness(2, 1, 2, 1));

			FrameworkElementFactory headerPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
			headerPresenter.SetValue(ContentPresenter.ContentSourceProperty, "Header");
			headerPresenter.SetBinding(ContentPresenter.ContentProperty,
			    new System.Windows.Data.Binding("Header")
			    {
			        RelativeSource = new System.Windows.Data.RelativeSource(
			            System.Windows.Data.RelativeSourceMode.TemplatedParent)
			    });
			headerBorder.AppendChild(headerPresenter);
			outerStack.AppendChild(headerBorder);

			// Items host for child items — bind Visibility to IsExpanded
			FrameworkElementFactory itemsHost = new FrameworkElementFactory(typeof(ItemsPresenter));
			itemsHost.SetValue(FrameworkElement.MarginProperty, new Thickness(16, 0, 0, 0));
			itemsHost.SetBinding(UIElement.VisibilityProperty,
			    new System.Windows.Data.Binding("IsExpanded")
			    {
			        RelativeSource = new System.Windows.Data.RelativeSource(
			            System.Windows.Data.RelativeSourceMode.TemplatedParent),
			        Converter = new System.Windows.Controls.BooleanToVisibilityConverter()
			    });
			outerStack.AppendChild(itemsHost);

			treeItemTemplate.VisualTree = outerStack;

			// Trigger: IsSelected → set Background + Foreground on the TreeViewItem itself
			// The Border binds to Background via TemplatedParent, so it picks it up automatically
			Trigger treeSelTrigger = new Trigger();
			treeSelTrigger.Property = TreeViewItem.IsSelectedProperty;
			treeSelTrigger.Value = true;
			treeSelTrigger.Setters.Add(new Setter(TreeViewItem.BackgroundProperty,
			    new SolidColorBrush(theme.TreeSelectedBackground)));
			treeSelTrigger.Setters.Add(new Setter(TreeViewItem.ForegroundProperty,
			    new SolidColorBrush(theme.TreeSelectedForeground)));
			treeItemTemplate.Triggers.Add(treeSelTrigger);

			treeItemStyle.Setters.Add(new Setter(TreeViewItem.TemplateProperty, treeItemTemplate));
			categoryTree.ItemContainerStyle = treeItemStyle;

			// ---- ListView ----
			itemList.Background = new SolidColorBrush(theme.ListBackground);
			itemList.Foreground = new SolidColorBrush(theme.ListForeground);

			// ListView item style with alternate row colors and hover
			Style listItemStyle = new Style(typeof(ListViewItem));
			listItemStyle.Setters.Add(new Setter(ListViewItem.ForegroundProperty, new SolidColorBrush(theme.ListForeground)));
			listItemStyle.Setters.Add(new Setter(ListViewItem.BackgroundProperty, Brushes.Transparent));
			listItemStyle.Setters.Add(new Setter(ListViewItem.BorderBrushProperty, Brushes.Transparent));
			listItemStyle.Setters.Add(new Setter(ListViewItem.BorderThicknessProperty, new Thickness(0)));

			// Custom template so we control Background rendering fully
			ControlTemplate listItemTemplate = new ControlTemplate(typeof(ListViewItem));
			FrameworkElementFactory listBorder = new FrameworkElementFactory(typeof(Border));
			listBorder.SetValue(FrameworkElement.NameProperty, "ItemBorder");
			listBorder.SetBinding(Border.BackgroundProperty,
			    new System.Windows.Data.Binding("Background") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
			listBorder.SetValue(Border.PaddingProperty, new Thickness(2, 0, 2, 0));

			FrameworkElementFactory gridRowPresenter = new FrameworkElementFactory(typeof(GridViewRowPresenter));
			gridRowPresenter.SetBinding(GridViewRowPresenter.ContentProperty,
			    new System.Windows.Data.Binding("."));
			gridRowPresenter.SetBinding(GridViewRowPresenter.ColumnsProperty,
			    new System.Windows.Data.Binding("View.Columns") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.FindAncestor, typeof(ListView), 1) });
			listBorder.AppendChild(gridRowPresenter);

			listItemTemplate.VisualTree = listBorder;

			// Selected trigger
			Trigger listSelTrigger = new Trigger();
			listSelTrigger.Property = ListViewItem.IsSelectedProperty;
			listSelTrigger.Value = true;
			listSelTrigger.Setters.Add(new Setter(ListViewItem.BackgroundProperty, new SolidColorBrush(theme.ListSelectedBackground)));
			listSelTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, new SolidColorBrush(theme.ListSelectedForeground)));
			listItemTemplate.Triggers.Add(listSelTrigger);

			// Hover trigger (only when not selected)
			MultiTrigger hoverTrigger = new MultiTrigger();
			hoverTrigger.Conditions.Add(new Condition(ListViewItem.IsMouseOverProperty, true));
			hoverTrigger.Conditions.Add(new Condition(ListViewItem.IsSelectedProperty, false));
			hoverTrigger.Setters.Add(new Setter(ListViewItem.BackgroundProperty, new SolidColorBrush(theme.ListHoverBackground)));
			hoverTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, new SolidColorBrush(theme.ListHoverForeground)));
			listItemTemplate.Triggers.Add(hoverTrigger);

			listItemStyle.Setters.Add(new Setter(ListViewItem.TemplateProperty, listItemTemplate));

			itemList.ItemContainerStyle = listItemStyle;

			// ListView column headers via style
			Style headerStyle = new Style(typeof(GridViewColumnHeader));
			headerStyle.Setters.Add(new Setter(GridViewColumnHeader.BackgroundProperty, new SolidColorBrush(theme.ListHeaderBackground)));
			headerStyle.Setters.Add(new Setter(GridViewColumnHeader.ForegroundProperty, new SolidColorBrush(theme.ListHeaderForeground)));
			headerStyle.Setters.Add(new Setter(GridViewColumnHeader.BorderBrushProperty, new SolidColorBrush(theme.ListHeaderBackground)));
			GridView gv = itemList.View as GridView;
			if (gv != null)
			    gv.ColumnHeaderContainerStyle = headerStyle;

		    // ---- Track Settings Panel ----
		    ApplyThemeToPanel(trackSettingsPanel, theme);

			// Theme the Track Settings toolbar
			SolidColorBrush tsTbBg = new SolidColorBrush(theme.ToolbarBackground);
			SolidColorBrush tsTbFg = new SolidColorBrush(theme.ToolbarForeground);
			foreach (var child in trackSettingsPanel.Children)
			{
			    if (child is ToolBarTray tbt)
			        ApplyThemeToToolBarTray(tbt, tsTbBg, tsTbFg);
			}

		    // ---- Tabs ----
		    SolidColorBrush tabBg = new SolidColorBrush(theme.TabBackground);
		    SolidColorBrush tabFg = new SolidColorBrush(theme.TabForeground);
		    SolidColorBrush tabSelBg = new SolidColorBrush(theme.TabSelectedBackground);
		    SolidColorBrush tabSelFg = new SolidColorBrush(theme.TabSelectedForeground);
		    rightTabs.Background = tabBg;
		    rightTabs.Foreground = tabFg;

		    Style tabItemStyle = new Style(typeof(TabItem));
		    tabItemStyle.Setters.Add(new Setter(TabItem.BackgroundProperty, tabBg));
		    tabItemStyle.Setters.Add(new Setter(TabItem.ForegroundProperty, tabFg));

		    // Custom ControlTemplate so Background actually renders on tab headers
		    ControlTemplate tabItemTemplate = new ControlTemplate(typeof(TabItem));

		    FrameworkElementFactory tabBorder = new FrameworkElementFactory(typeof(Border));
		    tabBorder.SetValue(FrameworkElement.NameProperty, "TabBorder");
		    tabBorder.SetBinding(Border.BackgroundProperty,
		        new System.Windows.Data.Binding("Background")
		        {
		            RelativeSource = new System.Windows.Data.RelativeSource(
		                System.Windows.Data.RelativeSourceMode.TemplatedParent)
		        });
		    tabBorder.SetValue(Border.BorderBrushProperty, new SolidColorBrush(theme.SplitterColor));
		    tabBorder.SetValue(Border.BorderThicknessProperty, new Thickness(1, 1, 1, 0));
		    tabBorder.SetValue(Border.PaddingProperty, new Thickness(8, 4, 8, 4));
		    tabBorder.SetValue(Border.MarginProperty, new Thickness(0, 0, 2, 0));
		    tabBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(4, 4, 0, 0));

		    FrameworkElementFactory tabContentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
		    tabContentPresenter.SetValue(ContentPresenter.ContentSourceProperty, "Header");
		    tabContentPresenter.SetBinding(ContentPresenter.ContentProperty,
		        new System.Windows.Data.Binding("Header")
		        {
		            RelativeSource = new System.Windows.Data.RelativeSource(
		                System.Windows.Data.RelativeSourceMode.TemplatedParent)
		        });
		    tabContentPresenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);

		    tabBorder.AppendChild(tabContentPresenter);
		    tabItemTemplate.VisualTree = tabBorder;

		    // Selected trigger
		    Trigger tabSelectedTrigger = new Trigger();
		    tabSelectedTrigger.Property = TabItem.IsSelectedProperty;
		    tabSelectedTrigger.Value = true;
		    tabSelectedTrigger.Setters.Add(new Setter(TabItem.BackgroundProperty, tabSelBg));
		    tabSelectedTrigger.Setters.Add(new Setter(TabItem.ForegroundProperty, tabSelFg));
		    tabSelectedTrigger.Setters.Add(new Setter(TabItem.MarginProperty, new Thickness(0, -2, 2, 0)));
		    tabItemTemplate.Triggers.Add(tabSelectedTrigger);

		    // Hover trigger (when not selected)
		    byte tabR = theme.TabBackground.R;
		    byte tabG = theme.TabBackground.G;
		    byte tabB = theme.TabBackground.B;
		    Color tabHoverColor;
		    if ((tabR + tabG + tabB) / 3 < 128)
		        tabHoverColor = Color.FromRgb(
		            (byte)Math.Min(255, tabR + 25),
		            (byte)Math.Min(255, tabG + 25),
		            (byte)Math.Min(255, tabB + 25));
		    else
		        tabHoverColor = Color.FromRgb(
		            (byte)Math.Max(0, tabR - 25),
		            (byte)Math.Max(0, tabG - 25),
		            (byte)Math.Max(0, tabB - 25));

		    MultiTrigger tabHoverTrigger = new MultiTrigger();
		    tabHoverTrigger.Conditions.Add(new Condition(TabItem.IsMouseOverProperty, true));
		    tabHoverTrigger.Conditions.Add(new Condition(TabItem.IsSelectedProperty, false));
		    tabHoverTrigger.Setters.Add(new Setter(TabItem.BackgroundProperty, new SolidColorBrush(tabHoverColor)));
		    tabItemTemplate.Triggers.Add(tabHoverTrigger);

		    tabItemStyle.Setters.Add(new Setter(TabItem.TemplateProperty, tabItemTemplate));

		    foreach (TabItem ti in rightTabs.Items)
		    {
		        ti.Style = tabItemStyle;
		    }

		    // ---- Theme toolbars inside tab content panels ----
		    SolidColorBrush tabTbBg = new SolidColorBrush(theme.ToolbarBackground);
		    SolidColorBrush tabTbFg = new SolidColorBrush(theme.ToolbarForeground);
		    foreach (TabItem ti in rightTabs.Items)
		    {
		        if (ti.Content is DockPanel dp)
		        {
		            foreach (var child in dp.Children)
		            {
		                if (child is ToolBarTray tbt && tbt != toolBarTray)
		                {
		                    ApplyThemeToToolBarTray(tbt, tabTbBg, tabTbFg);
		                }
		            }
		        }
		    }

		    // ---- Source View ----
		    sourceView.Background = new SolidColorBrush(theme.SourceBackground);
		    sourceView.Foreground = new SolidColorBrush(theme.SourceTagColor);

			// ---- Buttons (app-wide) ----
			SolidColorBrush btnBg = new SolidColorBrush(theme.ButtonBackground);
			SolidColorBrush btnFg = new SolidColorBrush(theme.ButtonForeground);

			// Compute button hover color
			byte btnR = theme.ButtonBackground.R;
			byte btnG = theme.ButtonBackground.G;
			byte btnB = theme.ButtonBackground.B;
			Color btnHoverColor;
			if ((btnR + btnG + btnB) / 3 < 128)
			    btnHoverColor = Color.FromRgb(
			        (byte)Math.Min(255, btnR + 40),
			        (byte)Math.Min(255, btnG + 40),
			        (byte)Math.Min(255, btnB + 40));
			else
			    btnHoverColor = Color.FromRgb(
			        (byte)Math.Max(0, btnR - 40),
			        (byte)Math.Max(0, btnG - 40),
			        (byte)Math.Max(0, btnB - 40));
			SolidColorBrush btnHoverBrush = new SolidColorBrush(btnHoverColor);

			Style buttonStyle = new Style(typeof(Button));
			buttonStyle.Setters.Add(new Setter(Button.BackgroundProperty, btnBg));
			buttonStyle.Setters.Add(new Setter(Button.ForegroundProperty, btnFg));
			buttonStyle.Setters.Add(new Setter(Button.BorderBrushProperty, btnBg));

			// Custom template so Background actually renders
			FrameworkElementFactory borderFactory = new FrameworkElementFactory(typeof(Border));
			borderFactory.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
			borderFactory.SetBinding(Border.BorderBrushProperty, new System.Windows.Data.Binding("BorderBrush") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
			borderFactory.SetValue(Border.BorderThicknessProperty, new Thickness(1));
			borderFactory.SetBinding(Border.PaddingProperty, new System.Windows.Data.Binding("Padding") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
			borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));

			FrameworkElementFactory contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
			contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
			contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
			borderFactory.AppendChild(contentPresenter);

			ControlTemplate buttonTemplate = new ControlTemplate(typeof(Button));
			buttonTemplate.VisualTree = borderFactory;

			// Hover trigger for global buttons
			Trigger globalBtnHover = new Trigger();
			globalBtnHover.Property = UIElement.IsMouseOverProperty;
			globalBtnHover.Value = true;
			globalBtnHover.Setters.Add(new Setter(Button.BackgroundProperty, btnHoverBrush));
			globalBtnHover.Setters.Add(new Setter(Button.BorderBrushProperty, btnHoverBrush));
			buttonTemplate.Triggers.Add(globalBtnHover);

			buttonStyle.Setters.Add(new Setter(Button.TemplateProperty, buttonTemplate));

			this.Resources[typeof(Button)] = buttonStyle;

		    // ---- Scrollbar resource overrides ----
		    this.Resources[SystemColors.ScrollBarBrushKey] = new SolidColorBrush(theme.ScrollBarBackground);

		    // Re-display source with new theme colors
		    if (!string.IsNullOrEmpty(currentSource))
		    {
		        DisplaySource(currentSource);
		    }

			// ---- Scrollbars ----
			SolidColorBrush scrollBg = new SolidColorBrush(theme.ScrollBarBackground);
			SolidColorBrush scrollThumb = new SolidColorBrush(theme.ScrollBarThumb);

			// Thumb style — this themes the draggable part of all scrollbars
			Style thumbStyle = new Style(typeof(System.Windows.Controls.Primitives.Thumb));
			ControlTemplate thumbTemplate = new ControlTemplate(typeof(System.Windows.Controls.Primitives.Thumb));
			FrameworkElementFactory thumbBorder = new FrameworkElementFactory(typeof(Border));
			thumbBorder.SetValue(Border.BackgroundProperty, scrollThumb);
			thumbBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));
			thumbBorder.SetValue(Border.MarginProperty, new Thickness(2));
			thumbTemplate.VisualTree = thumbBorder;
			thumbStyle.Setters.Add(new Setter(System.Windows.Controls.Primitives.Thumb.TemplateProperty, thumbTemplate));
			this.Resources[typeof(System.Windows.Controls.Primitives.Thumb)] = thumbStyle;

			// ScrollBar background color via system resource
			this.Resources[SystemColors.ScrollBarBrushKey] = scrollBg;

			// ScrollBar style with background
			Style scrollBarStyle = new Style(typeof(ScrollBar));
			scrollBarStyle.Setters.Add(new Setter(ScrollBar.BackgroundProperty, scrollBg));
			this.Resources[typeof(ScrollBar)] = scrollBarStyle;

		    // Rebuild list columns to pick up new StatusToColorConverter brushes
		    RebuildColumns();
		    // Force status icons to refresh with new theme colors
			RefreshStatusIcons(itemList);

			// ---- Theme ComboBox dropdown arrows ----
			// Walk all ComboBoxes and style their toggle buttons after rendering
			this.Dispatcher.BeginInvoke(new Action(() =>
			{
			    ThemeAllComboBoxArrows(this, new SolidColorBrush(theme.ToolbarForeground));
			}), System.Windows.Threading.DispatcherPriority.Loaded);
		}

		private void ApplyThemeToMenuItem(MenuItem mi, SolidColorBrush bg, SolidColorBrush fg, bool isTopLevel = true)
		{
			Style miStyle = new Style(typeof(MenuItem));
			miStyle.Setters.Add(new Setter(MenuItem.BackgroundProperty, bg));
			miStyle.Setters.Add(new Setter(MenuItem.ForegroundProperty, fg));
			mi.Style = miStyle;
		    mi.Foreground = fg;

		    // Compute hover color
		    byte r = bg.Color.R;
		    byte g = bg.Color.G;
		    byte b = bg.Color.B;
		    Color hoverColor;
		    if ((r + g + b) / 3 < 128)
		        hoverColor = Color.FromRgb(
		            (byte)Math.Min(255, r + 30),
		            (byte)Math.Min(255, g + 30),
		            (byte)Math.Min(255, b + 30));
		    else
		        hoverColor = Color.FromRgb(
		            (byte)Math.Max(0, r - 30),
		            (byte)Math.Max(0, g - 30),
		            (byte)Math.Max(0, b - 30));
		    SolidColorBrush hoverBrush = new SolidColorBrush(hoverColor);

		    if (isTopLevel)
		        mi.Template = BuildMenuItemTemplate(bg, fg, hoverBrush);
		    else
		        mi.Template = BuildSubmenuItemTemplate(bg, fg, hoverBrush);

		    if (mi.HasItems)
		    {
		        mi.SubmenuOpened -= MenuItem_SubmenuOpened;
		        mi.SubmenuOpened += MenuItem_SubmenuOpened;
		        mi.Tag = new SolidColorBrush[] { bg, fg };
		    }

		    foreach (var sub in mi.Items)
		    {
		        if (sub is MenuItem subMi)
		            ApplyThemeToMenuItem(subMi, bg, fg, false);
		        else if (sub is Separator sep)
		        {
		            sep.Margin = new Thickness(4, 2, 4, 2);
		            sep.Background = fg;
		        }
		    }
		}

		private void ApplyThemeToToolBarTray(ToolBarTray tray, SolidColorBrush bg, SolidColorBrush fg)
		{
		    tray.Background = bg;
		    foreach (ToolBar tb in tray.ToolBars)
		    {
		        tb.Background = bg;
		        tb.Foreground = fg;

		        // Compute hover color
		        byte r = bg.Color.R;
		        byte g = bg.Color.G;
		        byte b = bg.Color.B;
		        Color hoverColor;
		        if ((r + g + b) / 3 < 128)
		            hoverColor = Color.FromRgb(
		                (byte)Math.Min(255, r + 40),
		                (byte)Math.Min(255, g + 40),
		                (byte)Math.Min(255, b + 40));
		        else
		            hoverColor = Color.FromRgb(
		                (byte)Math.Max(0, r - 40),
		                (byte)Math.Max(0, g - 40),
		                (byte)Math.Max(0, b - 40));
		        SolidColorBrush hoverBrush = new SolidColorBrush(hoverColor);

		        // Build a toolbar-specific button style with hover effect
		        Style tbBtnStyle = new Style(typeof(Button));
		        tbBtnStyle.Setters.Add(new Setter(Button.ForegroundProperty, fg));
		        tbBtnStyle.Setters.Add(new Setter(Button.BackgroundProperty, bg));
		        tbBtnStyle.Setters.Add(new Setter(Button.BorderBrushProperty, bg));

		        FrameworkElementFactory btnBorder = new FrameworkElementFactory(typeof(Border));
		        btnBorder.SetValue(FrameworkElement.NameProperty, "TbBtnBorder");
		        btnBorder.SetBinding(Border.BackgroundProperty, new System.Windows.Data.Binding("Background") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
		        btnBorder.SetBinding(Border.BorderBrushProperty, new System.Windows.Data.Binding("BorderBrush") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
		        btnBorder.SetValue(Border.BorderThicknessProperty, new Thickness(1));
		        btnBorder.SetBinding(Border.PaddingProperty, new System.Windows.Data.Binding("Padding") { RelativeSource = new System.Windows.Data.RelativeSource(System.Windows.Data.RelativeSourceMode.TemplatedParent) });
		        btnBorder.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));

		        FrameworkElementFactory btnContent = new FrameworkElementFactory(typeof(ContentPresenter));
		        btnContent.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
		        btnContent.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
		        btnContent.SetValue(TextElement.ForegroundProperty, fg);
		        btnBorder.AppendChild(btnContent);

		        ControlTemplate btnTemplate = new ControlTemplate(typeof(Button));
		        btnTemplate.VisualTree = btnBorder;

		        // Hover trigger
		        Trigger btnHoverTrigger = new Trigger();
		        btnHoverTrigger.Property = UIElement.IsMouseOverProperty;
		        btnHoverTrigger.Value = true;
		        btnHoverTrigger.Setters.Add(new Setter(Button.BackgroundProperty, hoverBrush));
		        btnHoverTrigger.Setters.Add(new Setter(Button.BorderBrushProperty, hoverBrush));
		        btnTemplate.Triggers.Add(btnHoverTrigger);

		        tbBtnStyle.Setters.Add(new Setter(Button.TemplateProperty, btnTemplate));

		        // Apply to all items
		        foreach (var item in tb.Items)
		        {
		            if (item is Button btn)
		                btn.Style = tbBtnStyle;
		            else if (item is TextBlock txt)
		                txt.Foreground = fg;
		        }

		        // Override the ToolBar template
		        ControlTemplate tbTemplate = new ControlTemplate(typeof(ToolBar));
		        FrameworkElementFactory tbBorder = new FrameworkElementFactory(typeof(Border));
		        tbBorder.SetBinding(Border.BackgroundProperty,
		            new System.Windows.Data.Binding("Background")
		            {
		                RelativeSource = new System.Windows.Data.RelativeSource(
		                    System.Windows.Data.RelativeSourceMode.TemplatedParent)
		            });
		        tbBorder.SetValue(Border.PaddingProperty, new Thickness(2));

		        FrameworkElementFactory tbPanel = new FrameworkElementFactory(typeof(ToolBarPanel));
		        tbPanel.SetValue(ToolBarPanel.IsItemsHostProperty, true);
		        tbBorder.AppendChild(tbPanel);

		        tbTemplate.VisualTree = tbBorder;
		        tb.Template = tbTemplate;
		    }
		}

		private ControlTemplate BuildMenuItemTemplate(SolidColorBrush bg, SolidColorBrush fg, SolidColorBrush hoverBg)
		{
		    ControlTemplate template = new ControlTemplate(typeof(MenuItem));

		    FrameworkElementFactory outerGrid = new FrameworkElementFactory(typeof(Grid));

		    // Root border — binds background to MenuItem.Background
		    FrameworkElementFactory rootBorder = new FrameworkElementFactory(typeof(Border));
		    rootBorder.SetBinding(Border.BackgroundProperty,
		        new System.Windows.Data.Binding("Background")
		        {
		            RelativeSource = new System.Windows.Data.RelativeSource(
		                System.Windows.Data.RelativeSourceMode.TemplatedParent)
		        });
		    rootBorder.SetValue(Border.BorderBrushProperty, Brushes.Transparent);
		    rootBorder.SetValue(Border.BorderThicknessProperty, new Thickness(0));
		    rootBorder.SetValue(Border.PaddingProperty, new Thickness(6, 4, 6, 4));

		    FrameworkElementFactory contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
		    contentPresenter.SetValue(ContentPresenter.ContentSourceProperty, "Header");
		    contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
		    contentPresenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);

		    rootBorder.AppendChild(contentPresenter);
		    outerGrid.AppendChild(rootBorder);

		    // Popup for submenu items
		    FrameworkElementFactory popup = new FrameworkElementFactory(typeof(System.Windows.Controls.Primitives.Popup));
		    popup.SetValue(System.Windows.Controls.Primitives.Popup.PlacementProperty,
		        System.Windows.Controls.Primitives.PlacementMode.Bottom);
		    popup.SetBinding(System.Windows.Controls.Primitives.Popup.IsOpenProperty,
		        new System.Windows.Data.Binding("IsSubmenuOpen")
		        {
		            RelativeSource = new System.Windows.Data.RelativeSource(
		                System.Windows.Data.RelativeSourceMode.TemplatedParent)
		        });
		    popup.SetValue(System.Windows.Controls.Primitives.Popup.AllowsTransparencyProperty, true);

		    FrameworkElementFactory popupBorder = new FrameworkElementFactory(typeof(Border));
		    popupBorder.SetValue(Border.BackgroundProperty, bg);
		    popupBorder.SetValue(Border.BorderBrushProperty, fg);
		    popupBorder.SetValue(Border.BorderThicknessProperty, new Thickness(1));
		    popupBorder.SetValue(Border.PaddingProperty, new Thickness(0, 2, 0, 2));

		    FrameworkElementFactory itemsPresenter = new FrameworkElementFactory(typeof(ItemsPresenter));
		    popupBorder.AppendChild(itemsPresenter);
		    popup.AppendChild(popupBorder);

		    outerGrid.AppendChild(popup);

		    template.VisualTree = outerGrid;

		    // Hover trigger — sets MenuItem.Background, border picks it up via binding
		    Trigger hoverTrigger = new Trigger();
		    hoverTrigger.Property = MenuItem.IsHighlightedProperty;
		    hoverTrigger.Value = true;
		    hoverTrigger.Setters.Add(new Setter(MenuItem.BackgroundProperty, hoverBg));
		    template.Triggers.Add(hoverTrigger);

		    return template;
		}

		private ControlTemplate BuildSubmenuItemTemplate(SolidColorBrush bg, SolidColorBrush fg, SolidColorBrush hoverBg)
		{
		    ControlTemplate template = new ControlTemplate(typeof(MenuItem));

		    FrameworkElementFactory rootBorder = new FrameworkElementFactory(typeof(Border));
		    rootBorder.SetBinding(Border.BackgroundProperty,
		        new System.Windows.Data.Binding("Background")
		        {
		            RelativeSource = new System.Windows.Data.RelativeSource(
		                System.Windows.Data.RelativeSourceMode.TemplatedParent)
		        });
		    rootBorder.SetValue(Border.BorderBrushProperty, Brushes.Transparent);
		    rootBorder.SetValue(Border.BorderThicknessProperty, new Thickness(0));
		    rootBorder.SetValue(Border.PaddingProperty, new Thickness(8, 4, 8, 4));

		    FrameworkElementFactory contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
		    contentPresenter.SetValue(ContentPresenter.ContentSourceProperty, "Header");
		    contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
		    contentPresenter.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);

		    rootBorder.AppendChild(contentPresenter);
		    template.VisualTree = rootBorder;

		    // Hover trigger
		    Trigger hoverTrigger = new Trigger();
		    hoverTrigger.Property = MenuItem.IsHighlightedProperty;
		    hoverTrigger.Value = true;
		    hoverTrigger.Setters.Add(new Setter(MenuItem.BackgroundProperty, hoverBg));
		    template.Triggers.Add(hoverTrigger);

		    return template;
		}

		private void MenuItem_SubmenuOpened(object sender, RoutedEventArgs e)
		{
		    MenuItem mi = sender as MenuItem;
		    if (mi == null) return;

		    SolidColorBrush[] brushes = mi.Tag as SolidColorBrush[];
		    if (brushes == null || brushes.Length < 2) return;

		    // Re-apply templates to child items when submenu opens
		    foreach (var sub in mi.Items)
		    {
		        if (sub is MenuItem subMi)
		            ApplyThemeToMenuItem(subMi, brushes[0], brushes[1], false);
		    }
		}

		private void ApplyStatusBarLegendColors(StatusBar sb, ThemeSettings theme)
		{
		    foreach (var item in sb.Items)
		    {
		        if (item is StatusBarItem sbi)
		        {
		            if (sbi.Content is StackPanel legendPanel && legendPanel.Orientation == Orientation.Horizontal)
		            {
		                foreach (var child in legendPanel.Children)
		                {
		                    if (child is TextBlock tb && tb.FontFamily != null &&
		                        tb.FontFamily.Source == "Segoe Fluent Icons")
		                    {
		                        int idx = legendPanel.Children.IndexOf(tb);
		                        if (idx + 1 < legendPanel.Children.Count &&
		                            legendPanel.Children[idx + 1] is TextBlock labelTb)
		                        {
		                            string label = labelTb.Text.Trim().ToLower();
		                            switch (label)
		                            {
		                                case "unchanged":
		                                    tb.Foreground = new SolidColorBrush(theme.StatusUnchanged);
		                                    break;
		                                case "changed":
		                                    tb.Foreground = new SolidColorBrush(theme.StatusChanged);
		                                    break;
		                                case "error":
		                                    tb.Foreground = new SolidColorBrush(theme.StatusError);
		                                    break;
		                                case "new":
		                                    tb.Foreground = new SolidColorBrush(theme.StatusNew);
		                                    break;
		                                case "unchecked":
		                                    tb.Foreground = new SolidColorBrush(theme.StatusUnchecked);
		                                    break;
		                            }
		                            labelTb.Foreground = new SolidColorBrush(theme.StatusBarForeground);
		                        }
		                    }
		                }
		            }
		        }
		    }
		}

		private void ApplyThemeToElement<T>(Panel parent, Action<T> applyAction) where T : class
		{
		    foreach (var child in parent.Children)
		    {
		        if (child is T target)
		            applyAction(target);
		    }
		}

		private void ApplyThemeToPanel(DockPanel panel, ThemeSettings theme)
		{
		    panel.Background = new SolidColorBrush(theme.PanelBackground);
		    SolidColorBrush labelBrush = new SolidColorBrush(theme.PanelForeground);
		    SolidColorBrush tbBg = new SolidColorBrush(theme.TextBoxBackground);
		    SolidColorBrush tbFg = new SolidColorBrush(theme.TextBoxForeground);

		    ApplyThemeToPanelRecursive(panel, labelBrush, tbBg, tbFg);
		}
		
		private void ApplyThemeToPanelRecursive(DependencyObject parent, SolidColorBrush labelBrush, SolidColorBrush tbBg, SolidColorBrush tbFg)
		{
		    if (parent == null) return;

		    if (parent is TextBlock tb)
		    {
		        if (tb != startPositionText && tb != stopPositionText && tb != editLatestVersion &&
		            tb != statusFile && tb != findCount)
		        {
		            tb.Foreground = labelBrush;
		        }
		    }
		    else if (parent is TextBox txt)
		    {
		        txt.Background = tbBg;
		        txt.Foreground = tbFg;
		    }

		    // Walk logical children (works even before visual tree is ready)
		    if (parent is Panel p)
		    {
		        foreach (object child in p.Children)
		        {
		            if (child is DependencyObject d)
		                ApplyThemeToPanelRecursive(d, labelBrush, tbBg, tbFg);
		        }
		    }
		    else if (parent is Decorator dec && dec.Child != null)
		    {
		        ApplyThemeToPanelRecursive(dec.Child, labelBrush, tbBg, tbFg);
		    }
		    else if (parent is ContentControl cc && cc.Content is DependencyObject cd)
		    {
		        ApplyThemeToPanelRecursive(cd, labelBrush, tbBg, tbFg);
		    }
		    else if (parent is ItemsControl ic)
		    {
		        foreach (object item in ic.Items)
		        {
		            if (item is DependencyObject d)
		                ApplyThemeToPanelRecursive(d, labelBrush, tbBg, tbFg);
		        }
		    }

		    // Also try visual tree if loaded
		    int count = VisualTreeHelper.GetChildrenCount(parent);
		    for (int i = 0; i < count; i++)
		    {
		        DependencyObject child = VisualTreeHelper.GetChild(parent, i);
		        ApplyThemeToPanelRecursive(child, labelBrush, tbBg, tbFg);
		    }
		}

		private void RebuildColumns()
		{
		    if (itemList == null) return;

		    GridView gridView = itemList.View as GridView;
		    if (gridView == null) return;

		    // Remember current items and selection
		    int selectedIdx = itemList.SelectedIndex;

		    // Clear and rebuild all columns
		    gridView.Columns.Clear();

		    // Checkbox column
		    GridViewColumn checkCol = new GridViewColumn();
		    checkCol.Width = 30;
		    FrameworkElementFactory headerCheckbox = new FrameworkElementFactory(typeof(CheckBox));
		    headerCheckbox.AddHandler(CheckBox.ClickEvent, new RoutedEventHandler(SelectAll_Click));
		    checkCol.Header = new GridViewColumnHeader();
		    checkCol.HeaderTemplate = new DataTemplate();
		    checkCol.HeaderTemplate.VisualTree = headerCheckbox;
		    FrameworkElementFactory cellCheckbox = new FrameworkElementFactory(typeof(CheckBox));
		    cellCheckbox.SetBinding(CheckBox.IsCheckedProperty,
		        new System.Windows.Data.Binding("IsSelected"));
		    cellCheckbox.SetValue(CheckBox.HorizontalAlignmentProperty, HorizontalAlignment.Center);
		    checkCol.CellTemplate = new DataTemplate();
		    checkCol.CellTemplate.VisualTree = cellCheckbox;
		    gridView.Columns.Add(checkCol);

		    // Status column
		    GridViewColumn statusCol = new GridViewColumn();
		    statusCol.Header = "St";
		    statusCol.Width = 30;
		    statusCol.CellTemplate = CreateStatusCellTemplate();
		    gridView.Columns.Add(statusCol);

		    // Data columns
		    List<ColumnSetting> colSettings = windowSettings.ColumnSettings;
		    if (colSettings.Count == 0)
		        colSettings = WindowSettings.GetDefaultColumns();
		    colSettings.Sort((a, b) => a.DisplayIndex.CompareTo(b.DisplayIndex));

		    foreach (ColumnSetting cs in colSettings)
		    {
		        if (cs.Visible)
		        {
		            string statusBinding = GetStatusBindingForColumn(cs.Binding);
		            if (statusBinding != null)
		                gridView.Columns.Add(MakeColumn(cs.Header, cs.Binding, cs.Width, statusBinding));
		            else
		                gridView.Columns.Add(MakeColumn(cs.Header, cs.Binding, cs.Width));
		        }
		    }

		    // Re-apply header style
		    Style headerStyle = new Style(typeof(GridViewColumnHeader));
		    headerStyle.Setters.Add(new Setter(GridViewColumnHeader.BackgroundProperty, new SolidColorBrush(currentTheme.ListHeaderBackground)));
		    headerStyle.Setters.Add(new Setter(GridViewColumnHeader.ForegroundProperty, new SolidColorBrush(currentTheme.ListHeaderForeground)));
		    headerStyle.Setters.Add(new Setter(GridViewColumnHeader.BorderBrushProperty, new SolidColorBrush(currentTheme.ListHeaderBackground)));
		    gridView.ColumnHeaderContainerStyle = headerStyle;

		    // Restore selection
		    suppressAutoDownload = true;
		    itemList.ItemsSource = null;
		    itemList.ItemsSource = currentItems;
		    if (selectedIdx >= 0 && selectedIdx < currentItems.Count)
		        itemList.SelectedIndex = selectedIdx;
		    suppressAutoDownload = false;
		}

		private void ApplyThemeToSplitters(Grid grid, SolidColorBrush brush)
		{
		    foreach (var child in grid.Children)
		    {
		        if (child is GridSplitter gs)
		        {
		            gs.Background = brush;
		        }
		        else if (child is Grid subGrid)
		        {
		            ApplyThemeToSplitters(subGrid, brush);
		        }
		    }
		}


		private void ApplyThemeToPanelContent(DependencyObject obj, ThemeSettings theme)
		{
		    if (obj == null) return;

		    SolidColorBrush panelFg = new SolidColorBrush(theme.PanelForeground);
		    SolidColorBrush tbBg = new SolidColorBrush(theme.TextBoxBackground);
		    SolidColorBrush tbFg = new SolidColorBrush(theme.TextBoxForeground);

		    if (obj is TextBlock tb)
		    {
		        // Don't override special colored TextBlocks (like startPositionText, stopPositionText)
		        if (tb != startPositionText && tb != stopPositionText && tb != editLatestVersion)
		        {
		            tb.Foreground = panelFg;
		        }
		    }
		    else if (obj is TextBox txt)
		    {
		        txt.Background = tbBg;
		        txt.Foreground = tbFg;
		    }
		    else if (obj is Panel p)
		    {
		        foreach (var child in p.Children)
		        {
		            ApplyThemeToPanelContent(child as DependencyObject, theme);
		        }
		    }
		    else if (obj is Decorator d && d.Child != null)
		    {
		        ApplyThemeToPanelContent(d.Child, theme);
		    }
		    else if (obj is ContentControl cc && cc.Content is DependencyObject content)
		    {
		        ApplyThemeToPanelContent(content, theme);
		    }
		}


		private void ApplyForegroundRecursive(DependencyObject parent, SolidColorBrush panelFg, SolidColorBrush labelFg, ThemeSettings theme)
		{
		    int count = VisualTreeHelper.GetChildrenCount(parent);
		    // If not yet loaded, walk logical tree instead
		    if (count == 0 && parent is Panel p)
		    {
		        foreach (var child in p.Children)
		        {
		            if (child is TextBlock tb)
		            {
		                // Don't override status position text or latest version display
		                if (tb != startPositionText && tb != stopPositionText && tb != editLatestVersion && tb != statusFile && tb != findCount)
		                    tb.Foreground = labelFg;
		            }
		            if (child is DockPanel dp)
		                ApplyForegroundRecursive(dp, panelFg, labelFg, theme);
		            if (child is StackPanel sp)
		                ApplyForegroundRecursive(sp, panelFg, labelFg, theme);
		            if (child is ScrollViewer sv && sv.Content is DependencyObject svContent)
		                ApplyForegroundRecursive(svContent, panelFg, labelFg, theme);
		        }
		    }
		}

		private void ThemeMenuItemRecursive(MenuItem menuItem, ThemeSettings theme)
		{
		    menuItem.Background = new SolidColorBrush(theme.MenuBackground);
		    menuItem.Foreground = new SolidColorBrush(theme.MenuForeground);

		    foreach (var sub in menuItem.Items)
		    {
		        if (sub is MenuItem subItem)
		        {
		            ThemeMenuItemRecursive(subItem, theme);
		        }
		    }
		}

		private void ThemeStatusBarLegend(StatusBar statusBar, ThemeSettings theme)
		{
		    // Walk through the status bar to find the legend icons and update their colors
		    // The legend is in a StackPanel inside a StatusBarItem (HorizontalAlignment=Right)
		    foreach (var item in statusBar.Items)
		    {
		        if (item is StatusBarItem sbi)
		        {
		            if (sbi.Content is StackPanel sp)
		            {
		                foreach (var child in sp.Children)
		                {
		                    if (child is TextBlock tb)
		                    {
		                        // Check if this is an icon (Segoe Fluent Icons) or a label
		                        if (tb.FontFamily != null && tb.FontFamily.Source != null &&
		                            tb.FontFamily.Source.Contains("Segoe Fluent Icons"))
		                        {
		                            // It's a status icon — match by current color to figure out which one
		                            string text = tb.Text;
		                            // Find the next sibling TextBlock to identify which status this is
		                        }
		                        else
		                        {
		                            // It's a label text — update foreground
		                            string label = tb.Text.Trim();
		                            tb.Foreground = new SolidColorBrush(theme.StatusBarForeground);
		                        }
		                    }
		                }

		                // Now do a second pass to set icon colors by pairing icon+label
		                TextBlock pendingIcon = null;
		                foreach (var child in sp.Children)
		                {
		                    if (child is TextBlock tb)
		                    {
		                        if (tb.FontFamily != null && tb.FontFamily.Source != null &&
		                            tb.FontFamily.Source.Contains("Segoe Fluent Icons"))
		                        {
		                            pendingIcon = tb;
		                        }
		                        else if (pendingIcon != null)
		                        {
		                            string label = tb.Text.Trim();
		                            switch (label)
		                            {
		                                case "Unchanged":
		                                    pendingIcon.Foreground = new SolidColorBrush(theme.StatusUnchanged);
		                                    break;
		                                case "Changed":
		                                    pendingIcon.Foreground = new SolidColorBrush(theme.StatusChanged);
		                                    break;
		                                case "Error":
		                                    pendingIcon.Foreground = new SolidColorBrush(theme.StatusError);
		                                    break;
		                                case "New":
		                                    pendingIcon.Foreground = new SolidColorBrush(theme.StatusNew);
		                                    break;
		                                case "Unchecked":
		                                    pendingIcon.Foreground = new SolidColorBrush(theme.StatusUnchecked);
		                                    break;
		                            }
		                            pendingIcon = null;
		                        }
		                    }
		                }
		            }
		        }
		    }
		}

		private void RefreshStatusIcons(ListView list)
		{
		    for (int i = 0; i < list.Items.Count; i++)
		    {
		        ListViewItem container = list.ItemContainerGenerator.ContainerFromIndex(i) as ListViewItem;
		        if (container == null) continue;

		        // Find TextBlocks using Segoe Fluent Icons (the status icon)
		        FindAndRefreshStatusIcons(container);
		    }
		}

		private void FindAndRefreshStatusIcons(DependencyObject parent)
		{
		    int count = VisualTreeHelper.GetChildrenCount(parent);
		    for (int i = 0; i < count; i++)
		    {
		        DependencyObject child = VisualTreeHelper.GetChild(parent, i);
		        if (child is TextBlock tb && tb.FontFamily != null &&
		            tb.FontFamily.Source == "Segoe Fluent Icons")
		        {
		            TrackItem item = tb.DataContext as TrackItem;
		            if (item != null)
		            {
		                ApplyStatusColor(tb, item.TrackStatus);
		            }
		        }
		        FindAndRefreshStatusIcons(child);
		    }
		}

        // ============================
        // Toolbar Docking
        // ============================

        private void DockToolbar_Click(object sender, RoutedEventArgs e)
        {
            MenuItem menuItem = sender as MenuItem;
            if (menuItem == null) return;

            string position = menuItem.Tag as string;
            if (position != null)
            {
                DockToolbarTo(position);
            }
        }

        private void DockToolbarTo(string position)
        {
            DockPanel parent = toolBarTray.Parent as DockPanel;
            if (parent == null) return;

            parent.Children.Remove(toolBarTray);

            switch (position)
            {
                case "Top":
                    DockPanel.SetDock(toolBarTray, Dock.Top);
                    toolBarTray.Orientation = Orientation.Horizontal;
                    break;
                case "Bottom":
                    DockPanel.SetDock(toolBarTray, Dock.Bottom);
                    toolBarTray.Orientation = Orientation.Horizontal;
                    break;
                case "Left":
                    DockPanel.SetDock(toolBarTray, Dock.Left);
                    toolBarTray.Orientation = Orientation.Vertical;
                    break;
                case "Right":
                    DockPanel.SetDock(toolBarTray, Dock.Right);
                    toolBarTray.Orientation = Orientation.Vertical;
                    break;
            }

            parent.Children.Insert(1, toolBarTray);
            statusFile.Text = "Toolbar docked: " + position;
        }

        // ============================
        // Input Dialog
        // ============================

        private string ShowInputDialog(string title, string prompt)
        {
            Window dialog = new Window();
            dialog.Title = title;
            dialog.Width = 350;
            dialog.Height = 160;
            dialog.WindowStartupLocation = WindowStartupLocation.CenterOwner;
            dialog.Owner = this;
            dialog.ResizeMode = ResizeMode.NoResize;

            StackPanel panel = new StackPanel();
            panel.Margin = new Thickness(12);

            TextBlock label = new TextBlock();
            label.Text = prompt;
            label.Margin = new Thickness(0, 0, 0, 8);
            panel.Children.Add(label);

            TextBox input = new TextBox();
            input.Margin = new Thickness(0, 0, 0, 12);
            panel.Children.Add(input);

            StackPanel buttons = new StackPanel();
            buttons.Orientation = Orientation.Horizontal;
            buttons.HorizontalAlignment = HorizontalAlignment.Right;

            Button btnOK = new Button();
            btnOK.Content = "OK";
            btnOK.Width = 70;
            btnOK.Padding = new Thickness(4);
            btnOK.Margin = new Thickness(0, 0, 8, 0);
            btnOK.IsDefault = true;
            btnOK.Click += (s, ev) => { dialog.DialogResult = true; };
            buttons.Children.Add(btnOK);

            Button btnCancel = new Button();
            btnCancel.Content = "Cancel";
            btnCancel.Width = 70;
            btnCancel.Padding = new Thickness(4);
            btnCancel.IsCancel = true;
            buttons.Children.Add(btnCancel);

            panel.Children.Add(buttons);
            dialog.Content = panel;

            dialog.Loaded += (s, ev) => { input.Focus(); };

            if (dialog.ShowDialog() == true)
            {
                return input.Text.Trim();
            }

            return null;
        }

        // ============================
        // Menu Handlers
        // ============================

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        // ============================
        // Toolbar Handlers
        // ============================

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            SaveTrack_Click(sender, e);
        }

        private void MenuNewTrack_Click(object sender, RoutedEventArgs e)
        {
            TreeViewItem selectedCat = categoryTree.SelectedItem as TreeViewItem;
            if (selectedCat == null)
            {
                MessageBox.Show("Please select a category first.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // Clear fields so user starts fresh
            ClearTrackFields();

            // Prompt for name and create
            string name = ShowInputDialog("New Track", "Enter program name:");
            if (string.IsNullOrWhiteSpace(name))
                return;

            editName.Text = name;
            CreateNewTrack();
        }

        private void SaveTrackAs_Click(object sender, RoutedEventArgs e)
        {
            TrackItem selected = currentTrackItem;

            if (selected == null)
            {
                MessageBox.Show("No track selected.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            string newName = ShowInputDialog("Save Track As",
                "Enter new name for the track copy:");

            if (string.IsNullOrWhiteSpace(newName))
                return;

            string safeName = newName;
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(c.ToString(), "");
            }

            if (string.IsNullOrWhiteSpace(safeName))
            {
                MessageBox.Show("Invalid file name.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string dir = Path.GetDirectoryName(selected.FilePath);
            string newPath = Path.Combine(dir, safeName + ".track");

            if (File.Exists(newPath))
            {
                MessageBoxResult overwrite = MessageBox.Show(
                    "Track file '" + safeName + ".track' already exists.\n\n" +
                    "Do you want to overwrite it?",
                    "PAT v7", MessageBoxButton.YesNo, MessageBoxImage.Question);

                if (overwrite != MessageBoxResult.Yes)
                    return;
            }

            try
            {
                // Create a copy with current field values
                TrackItem copy = new TrackItem();
                copy.ProgramName = newName;
                copy.FilePath = newPath;
                copy.TrackURL = editTrackURL.Text;
                copy.StartString = editStartString.Text;
                copy.StopString = editStopString.Text;
                copy.DownloadURL = editDownloadURL.Text;
                copy.Version = editVersion.Text;
                copy.ReleaseDate = editReleaseDate.Text;
                copy.PublisherName = editPublisherName.Text;
                copy.SuiteName = editSuiteName.Text;
                copy.CreationDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                copy.SaveToFile(newPath);

                // Reload the category and select the new track
                string folderPath = currentCategoryPath;
                if (string.IsNullOrEmpty(folderPath))
                {
                    TreeViewItem selectedCat = categoryTree.SelectedItem as TreeViewItem;
                    if (selectedCat != null)
                        folderPath = (string)selectedCat.Tag;
                }

                if (!string.IsNullOrEmpty(folderPath))
                {
                    LoadTrackFiles(folderPath);
                }
                RefreshCategoryTree();

                // Select the new copy
                for (int i = 0; i < currentItems.Count; i++)
                {
                    if (currentItems[i].FilePath == newPath)
                    {
                        itemList.SelectedIndex = i;
                        break;
                    }
                }

                statusFile.Text = "Saved as: " + newName;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving track copy:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void InstallUBlock_Click(object sender, RoutedEventArgs e)
        {
            string ublockFolder = Path.Combine(extensionsPath, "uBlock0.chromium");

            statusFile.Text = "Downloading uBlock Origin...";
            statusProgress.IsIndeterminate = true;

            try
            {
                // Download latest uBlock Origin for Chromium from GitHub releases
                HttpClient client = new HttpClient();
                client.DefaultRequestHeaders.Add("User-Agent", UserAgent);

                // Get latest release info from GitHub API
                string apiUrl = "https://api.github.com/repos/gorhill/uBlock/releases/latest";
                string json = await client.GetStringAsync(apiUrl);

                // Find the chromium zip asset URL
                string assetUrl = null;
                string[] lines = json.Split(',');
                foreach (string line in lines)
                {
                    if (line.Contains("browser_download_url") && line.Contains("uBlock0") && line.Contains(".chromium.zip"))
                    {
                        int urlStart = line.IndexOf("https://");
                        int urlEnd = line.IndexOf("\"", urlStart);
                        if (urlStart >= 0 && urlEnd > urlStart)
                        {
                            assetUrl = line.Substring(urlStart, urlEnd - urlStart);
                        }
                        break;
                    }
                }

                if (assetUrl == null)
                {
                    MessageBox.Show("Could not find uBlock Origin Chromium release.\n" +
                        "You can manually download from:\nhttps://github.com/gorhill/uBlock/releases\n\n" +
                        "Extract the uBlock0.chromium folder to:\n" + extensionsPath,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                statusFile.Text = "Downloading: " + assetUrl;

                // Download the zip
                byte[] zipBytes = await client.GetByteArrayAsync(assetUrl);
                string zipPath = Path.Combine(extensionsPath, "ublock_chromium.zip");
                File.WriteAllBytes(zipPath, zipBytes);

                // Remove old folder if exists
                if (Directory.Exists(ublockFolder))
                    Directory.Delete(ublockFolder, true);

                // Extract
                System.IO.Compression.ZipFile.ExtractToDirectory(zipPath, extensionsPath);

                // Clean up zip
                File.Delete(zipPath);

                if (!Directory.Exists(ublockFolder))
                {
                    // Check if extracted to a different folder name
                    string[] dirs = Directory.GetDirectories(extensionsPath, "uBlock*");
                    if (dirs.Length > 0)
                    {
                        ublockFolder = dirs[0];
                    }
                }

                if (!File.Exists(Path.Combine(ublockFolder, "manifest.json")))
                {
                    MessageBox.Show("Download succeeded but manifest.json not found.\n" +
                        "Check: " + extensionsPath,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Switch to WebView tab so WebView2 can initialise
                rightTabs.SelectedIndex = 2;
                await System.Threading.Tasks.Task.Delay(100);

                // Install into WebView2
                await webView.EnsureCoreWebView2Async();
                var extension = await webView.CoreWebView2.Profile.AddBrowserExtensionAsync(ublockFolder);

                statusFile.Text = "uBlock Origin installed successfully! Extension ID: " + extension.Id;
                MessageBox.Show("uBlock Origin installed successfully!\n\n" +
                    "The extension will be active for all WebView navigation.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error installing uBlock Origin:\n" + ex.Message +
                    "\n\nYou can manually download from:\nhttps://github.com/gorhill/uBlock/releases\n\n" +
                    "Extract the uBlock0.chromium folder to:\n" + extensionsPath,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                statusFile.Text = "Extension install failed: " + ex.Message;
            }
            finally
            {
                statusProgress.IsIndeterminate = false;
                statusProgress.Value = 0;
            }
        }

        private async void ListExtensions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await webView.EnsureCoreWebView2Async();
                var extensions = await webView.CoreWebView2.Profile.GetBrowserExtensionsAsync();

                if (extensions.Count == 0)
                {
                    MessageBox.Show("No extensions installed.", "PAT v7",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Installed Extensions:");
                sb.AppendLine();

                foreach (var ext in extensions)
                {
                    sb.AppendLine("Name: " + ext.Name);
                    sb.AppendLine("ID: " + ext.Id);
                    sb.AppendLine("Enabled: " + ext.IsEnabled);
                    sb.AppendLine();
                }

                MessageBox.Show(sb.ToString(), "WebView2 Extensions",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error listing extensions:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ============================
        // Browser + Download
        // ============================

        /// <summary>
        /// Downloads page source using WebView2 as a fallback when HttpClient fails (e.g. Cloudflare).
        /// WebView2 is a real browser and can pass JavaScript challenges.
        /// </summary>
        private async System.Threading.Tasks.Task<string> DownloadViaWebView(string url)
        {
            try
            {
                await webView.EnsureCoreWebView2Async();

                var tcs = new System.Threading.Tasks.TaskCompletionSource<string>();
                int retries = 0;
                const int maxRetries = 3;

                EventHandler<Microsoft.Web.WebView2.Core.CoreWebView2NavigationCompletedEventArgs> handler = null;
                handler = async (s, e) =>
                {
                    try
                    {
                        // Wait a moment for any Cloudflare challenge to resolve
                        await System.Threading.Tasks.Task.Delay(2000);

                        // Check if we're still on a Cloudflare challenge page
                        string title = webView.CoreWebView2.DocumentTitle ?? "";
                        string currentUrl = webView.CoreWebView2.Source ?? "";

                        if ((title.Contains("Just a moment") || title.Contains("Checking your browser") ||
                             title.Contains("Attention Required")) && retries < maxRetries)
                        {
                            retries++;
                            statusFile.Text = "Cloudflare challenge detected, waiting... (attempt " +
                                retries + "/" + maxRetries + ")";
                            await System.Threading.Tasks.Task.Delay(3000);
                            return; // Wait for next NavigationCompleted
                        }

                        // Get the page source via JavaScript
                        string source = await webView.CoreWebView2.ExecuteScriptAsync(
                            "document.documentElement.outerHTML");

                        if (!string.IsNullOrEmpty(source) && source != "null")
                        {
                            // The result comes back as a JSON string, so unescape it
                            source = System.Text.Json.JsonSerializer.Deserialize<string>(source);
                        }

                        webView.CoreWebView2.NavigationCompleted -= handler;

                        if (!string.IsNullOrEmpty(source) && source.Length > 100)
                        {
                            tcs.TrySetResult(source);
                        }
                        else
                        {
                            tcs.TrySetResult(null);
                        }
                    }
                    catch (Exception)
                    {
                        webView.CoreWebView2.NavigationCompleted -= handler;
                        tcs.TrySetResult(null);
                    }
                };

                webView.CoreWebView2.NavigationCompleted += handler;
                webView.CoreWebView2.Navigate(url);

                // Timeout after 30 seconds
                var timeoutTask = System.Threading.Tasks.Task.Delay(30000);
                var completedTask = await System.Threading.Tasks.Task.WhenAny(tcs.Task, timeoutTask);

                if (completedTask == timeoutTask)
                {
                    webView.CoreWebView2.NavigationCompleted -= handler;
                    return null;
                }

                return await tcs.Task;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void SetupHttpHeaders(string url = null)
        {
            httpClient.DefaultRequestHeaders.Clear();
            httpClient.DefaultRequestHeaders.Add("User-Agent", UserAgent);
            httpClient.DefaultRequestHeaders.Add("Accept",
                "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
            httpClient.DefaultRequestHeaders.Add("Accept-Language", "en-US,en;q=0.5");
            httpClient.DefaultRequestHeaders.Add("Connection", "keep-alive");
            httpClient.DefaultRequestHeaders.Add("Upgrade-Insecure-Requests", "1");

            if (!string.IsNullOrEmpty(url))
            {
                try
                {
                    Uri uri = new Uri(url);
                    httpClient.DefaultRequestHeaders.Add("Referer", uri.GetLeftPart(UriPartial.Authority) + "/");
                }
                catch (Exception) { }
            }
        }

        private void OpenBrowser_Click(object sender, RoutedEventArgs e)
        {
            string url = editTrackURL.Text;

            if (!string.IsNullOrWhiteSpace(url))
            {
                try
                {
                    System.Diagnostics.Process.Start(
                        new System.Diagnostics.ProcessStartInfo(url)
                        { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error opening browser: " + ex.Message,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void DownloadPage_Click(object sender, RoutedEventArgs e)
        {
            string url = editTrackURL.Text;

            if (string.IsNullOrWhiteSpace(url))
            {
                MessageBox.Show("Enter a Track URL first.", "PAT v7",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!url.StartsWith("http://") && !url.StartsWith("https://"))
            {
                url = "https://" + url;
                editTrackURL.Text = url;
            }

            statusFile.Text = "Downloading: " + url;
            statusProgress.IsIndeterminate = true;

            try
            {
                SetupHttpHeaders(url);

                currentSource = await httpClient.GetStringAsync(url);
                long downloadBytes = System.Text.Encoding.UTF8.GetByteCount(currentSource);

                DisplaySource(currentSource);

                // Apply tracking check
                TrackItem selected = currentTrackItem;
                if (selected != null)
                {
                    selected.ProgramName = editName.Text;
                    selected.TrackURL = editTrackURL.Text;
                    selected.StartString = editStartString.Text;
                    selected.StopString = editStopString.Text;
                    selected.DownloadURL = editDownloadURL.Text;
                    selected.Version = editVersion.Text;

                    CheckResult result = selected.ApplyCheck(currentSource, downloadBytes);

                    selected.SaveToFile();

                    // Refresh list but preserve selection
                    int selectedIndex = currentItems.IndexOf(selected);
                    suppressAutoDownload = true;
                    itemList.ItemsSource = null;
                    itemList.ItemsSource = currentItems;
                    if (selectedIndex >= 0)
                        itemList.SelectedIndex = selectedIndex;
                    suppressAutoDownload = false;

                    // Update UI fields immediately
                    editVersion.Text = selected.Version;

                    // Only show latest version if we actually have start/stop strings
                    if (string.IsNullOrEmpty(selected.StartString) ||
                        string.IsNullOrEmpty(selected.StopString))
                    {
                        selected.LatestVersion = "";
                    }
                    UpdateVersionDisplay();
                    statusFile.Text = "Checked: " + result.ProgramName +
                        " — " + result.Status.ToUpper() + " — " + result.Note;
                    editReleaseDate.Text = selected.ReleaseDate;
                }
                else
                {
                    statusFile.Text = "Downloaded: " + url +
                        " (" + currentSource.Length.ToString("N0") + " chars)" +
                        " — Save the track first to apply checks";
                }

                // Stop progress bar before WebView navigation (which can take time)
                statusProgress.IsIndeterminate = false;
                statusProgress.Value = 0;

                // Navigate WebView
                try
                {
                    await webView.EnsureCoreWebView2Async();
                    webView.CoreWebView2.Navigate(url);
                }
                catch (Exception)
                {
                }
            }
            catch (Exception ex)
            {
                // HttpClient failed — try WebView2 fallback
                statusFile.Text = "HttpClient blocked (" + ex.Message + "), trying WebView2...";

                try
                {
                    string source = await DownloadViaWebView(url);
                    if (!string.IsNullOrEmpty(source))
                    {
                        currentSource = source;
                        long downloadBytes = System.Text.Encoding.UTF8.GetByteCount(currentSource);
                        DisplaySource(currentSource);

                        TrackItem selected = currentTrackItem;
                        if (selected != null)
                        {
                            selected.ProgramName = editName.Text;
                            selected.TrackURL = editTrackURL.Text;
                            selected.StartString = editStartString.Text;
                            selected.StopString = editStopString.Text;
                            selected.DownloadURL = editDownloadURL.Text;
                            selected.Version = editVersion.Text;

                            CheckResult result = selected.ApplyCheck(currentSource, downloadBytes);
                            selected.SaveToFile();

                            int selectedIndex = currentItems.IndexOf(selected);
                            suppressAutoDownload = true;
                            itemList.ItemsSource = null;
                            itemList.ItemsSource = currentItems;
                            if (selectedIndex >= 0)
                                itemList.SelectedIndex = selectedIndex;
                            suppressAutoDownload = false;

                            editVersion.Text = selected.Version;
                            if (string.IsNullOrEmpty(selected.StartString) ||
                                string.IsNullOrEmpty(selected.StopString))
                            {
                                selected.LatestVersion = "";
                            }
                            UpdateVersionDisplay();
                            statusFile.Text = "Checked (WebView2): " + result.ProgramName +
                                " — " + result.Status.ToUpper() + " — " + result.Note;
                            editReleaseDate.Text = selected.ReleaseDate;
                        }
                        else
                        {
                            statusFile.Text = "Downloaded via WebView2: " + url;
                        }
                    }
                    else
                    {
                        TrackItem selected = currentTrackItem;
                        if (selected != null)
                        {
                            selected.TrackStatus = "error";
                            selected.LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                            selected.SaveToFile();
                            itemList.ItemsSource = null;
                            itemList.ItemsSource = currentItems;
                        }
                        MessageBox.Show("Download failed:\n" + ex.Message +
                            "\n\nWebView2 fallback also failed.",
                            "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                        statusFile.Text = "Download failed";
                    }
                }
                catch (Exception ex2)
                {
                    TrackItem selected = currentTrackItem;
                    if (selected != null)
                    {
                        selected.TrackStatus = "error";
                        selected.LastChecked = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                        selected.SaveToFile();
                        itemList.ItemsSource = null;
                        itemList.ItemsSource = currentItems;
                    }
                    MessageBox.Show("Download failed:\n" + ex.Message +
                        "\n\nWebView2 fallback error:\n" + ex2.Message,
                        "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
                    statusFile.Text = "Download failed";
                }
            }
            finally
            {
                statusProgress.IsIndeterminate = false;
                statusProgress.Value = 0;
            }
        }

        private void DisplaySource(string source)
        {
            sourceView.Document.Blocks.Clear();

            if (string.IsNullOrEmpty(source))
                return;

            Paragraph para = new Paragraph();
            para.FontFamily = new FontFamily("Consolas");
            para.FontSize = sourceView.FontSize;

            string startStr = editStartString.Text ?? "";
            string stopStr = editStopString.Text ?? "";

            try
            {
                if (!string.IsNullOrEmpty(startStr) && !string.IsNullOrEmpty(stopStr))
                {
                    string normalizedSource = source.Replace("\r\n", "\n").Replace("\r", "\n");

                    // Find start match
                    int startIdx = -1;
                    int startLen = 0;
                    int stopIdx = -1;
                    int stopLen = 0;

                    bool startIsRegex = TrackItem.IsRegexPattern(startStr);
                    bool stopIsRegex = TrackItem.IsRegexPattern(stopStr);

                    if (startIsRegex)
                    {
                        var match = TrackItem.FindRegexMatch(normalizedSource, startStr, 0);
                        if (match != null && match.Success)
                        {
                            startIdx = match.Index;
                            startLen = match.Length;
                        }
                    }
                    else
                    {
                        string normalizedStart = startStr.Replace("\r\n", "\n").Replace("\r", "\n");
                        startIdx = normalizedSource.IndexOf(normalizedStart, StringComparison.Ordinal);
                        startLen = normalizedStart.Length;

                        if (startIdx < 0)
                        {
                            string collapsedSource = NormalizeWhitespace(normalizedSource);
                            string collapsedStart = NormalizeWhitespace(normalizedStart);
                            int cIdx = collapsedSource.IndexOf(collapsedStart, StringComparison.Ordinal);
                            if (cIdx >= 0)
                            {
                                startIdx = MapNormalizedToRaw(normalizedSource, cIdx);
                                int startEndRaw = MapNormalizedToRaw(normalizedSource,
                                    cIdx + collapsedStart.Length);
                                startLen = startEndRaw - startIdx;
                            }
                        }
                    }

                    // Find stop match
                    if (startIdx >= 0)
                    {
                        int searchAfter = startIdx + startLen;

                        if (stopIsRegex)
                        {
                            var match = TrackItem.FindRegexMatch(normalizedSource, stopStr, searchAfter);
                            if (match != null && match.Success)
                            {
                                stopIdx = match.Index;
                                stopLen = match.Length;
                            }
                        }
                        else
                        {
                            string normalizedStop = stopStr.Replace("\r\n", "\n").Replace("\r", "\n");
                            stopIdx = normalizedSource.IndexOf(normalizedStop, searchAfter,
                                StringComparison.Ordinal);
                            stopLen = normalizedStop.Length;

                            if (stopIdx < 0)
                            {
                                string collapsedSource = NormalizeWhitespace(normalizedSource);
                                string collapsedStop = NormalizeWhitespace(normalizedStop);
                                string collapsedBefore = NormalizeWhitespace(
                                    normalizedSource.Substring(0, searchAfter));
                                int collapsedSearchAfter = collapsedBefore.Length;

                                int cIdx = collapsedSource.IndexOf(collapsedStop, collapsedSearchAfter,
                                    StringComparison.Ordinal);
                                if (cIdx >= 0)
                                {
                                    stopIdx = MapNormalizedToRaw(normalizedSource, cIdx);
                                    int stopEndRaw = MapNormalizedToRaw(normalizedSource,
                                        cIdx + collapsedStop.Length);
                                    stopLen = stopEndRaw - stopIdx;
                                }
                            }
                        }
                    }

                    // Render with highlighting
                    if (startIdx >= 0 && stopIdx >= 0 &&
                        startLen > 0 && stopLen > 0 &&
                        startIdx + startLen <= normalizedSource.Length &&
                        stopIdx + stopLen <= normalizedSource.Length)
                    {
                        int startEnd = startIdx + startLen;
                        int stopEnd = stopIdx + stopLen;

                        if (startIdx > 0)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(0, startIdx)))
                                para.Inlines.Add(r);
                        }

                        Run startRun = new Run(normalizedSource.Substring(startIdx, startLen));
                        startRun.Foreground = new SolidColorBrush(currentTheme.SourceStartStringColor);
                        startRun.FontWeight = FontWeights.Bold;
                        para.Inlines.Add(startRun);

                        int betweenLength = stopIdx - startEnd;
                        if (betweenLength > 0)
                        {
                            Run betweenRun = new Run(normalizedSource.Substring(startEnd, betweenLength));
                            betweenRun.Foreground = new SolidColorBrush(currentTheme.SourceInfoStringColor);
                            betweenRun.FontWeight = FontWeights.Bold;
                            para.Inlines.Add(betweenRun);
                        }

                        Run stopRun = new Run(normalizedSource.Substring(stopIdx, stopLen));
                        stopRun.Foreground = new SolidColorBrush(currentTheme.SourceStopStringColor);
                        stopRun.FontWeight = FontWeights.Bold;
                        para.Inlines.Add(stopRun);

                        if (stopEnd < normalizedSource.Length)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(stopEnd)))
                                para.Inlines.Add(r);
                        }

                        startPositionText.Text = "pos: " + startIdx;
                        startPositionText.Foreground = Brushes.Green;
                        stopPositionText.Text = "pos: " + stopIdx;
                        stopPositionText.Foreground = Brushes.Green;

                        string trackBlock = normalizedSource.Substring(startIdx, stopEnd - startIdx);
                        string extractedVersion = TrackItem.ExtractVersion(trackBlock);
                        if (!string.IsNullOrEmpty(extractedVersion) && currentTrackItem != null)
                        {
                            currentTrackItem.LatestVersion = extractedVersion;
                            UpdateVersionDisplay();
                        }
                    }
                    else if (startIdx >= 0 && startLen > 0 &&
                             startIdx + startLen <= normalizedSource.Length)
                    {
                        int startEnd = startIdx + startLen;

                        if (startIdx > 0)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(0, startIdx)))
                                para.Inlines.Add(r);
                        }

                        Run startRun = new Run(normalizedSource.Substring(startIdx, startLen));
                        startRun.Foreground = new SolidColorBrush(currentTheme.SourceStartStringColor);
                        startRun.FontWeight = FontWeights.Bold;
                        para.Inlines.Add(startRun);

                        if (startEnd < normalizedSource.Length)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(startEnd)))
                                para.Inlines.Add(r);
                        }

                        startPositionText.Text = "pos: " + startIdx;
                        startPositionText.Foreground = Brushes.Green;
                        stopPositionText.Text = "NOT FOUND";
                        stopPositionText.Foreground = Brushes.Red;
                    }
                    else
                    {
                        foreach (Run r in CreateSyntaxHighlightedRuns(normalizedSource))
                            para.Inlines.Add(r);

                        startPositionText.Text = "NOT FOUND";
                        startPositionText.Foreground = Brushes.Red;
                        stopPositionText.Text = "";
                    }
                }
                else if (!string.IsNullOrEmpty(startStr))
                {
                    // Only start string provided — highlight just that
                    string normalizedSource = source.Replace("\r\n", "\n").Replace("\r", "\n");
                    int startIdx = -1;
                    int startLen = 0;

                    if (TrackItem.IsRegexPattern(startStr))
                    {
                        var match = TrackItem.FindRegexMatch(normalizedSource, startStr, 0);
                        if (match != null && match.Success)
                        {
                            startIdx = match.Index;
                            startLen = match.Length;
                        }
                    }
                    else
                    {
                        string normalizedStart = startStr.Replace("\r\n", "\n").Replace("\r", "\n");
                        startIdx = normalizedSource.IndexOf(normalizedStart, StringComparison.Ordinal);
                        startLen = normalizedStart.Length;

                        if (startIdx < 0)
                        {
                            string cs = NormalizeWhitespace(normalizedSource);
                            string cst = NormalizeWhitespace(normalizedStart);
                            int ci = cs.IndexOf(cst, StringComparison.Ordinal);
                            if (ci >= 0)
                            {
                                startIdx = MapNormalizedToRaw(normalizedSource, ci);
                                int endRaw = MapNormalizedToRaw(normalizedSource, ci + cst.Length);
                                startLen = endRaw - startIdx;
                            }
                        }
                    }

                    if (startIdx >= 0 && startLen > 0 && startIdx + startLen <= normalizedSource.Length)
                    {
                        if (startIdx > 0)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(0, startIdx)))
                                para.Inlines.Add(r);
                        }

                        Run startRun = new Run(normalizedSource.Substring(startIdx, startLen));
                        startRun.Foreground = new SolidColorBrush(currentTheme.SourceStartStringColor);
                        startRun.FontWeight = FontWeights.Bold;
                        para.Inlines.Add(startRun);

                        if (startIdx + startLen < normalizedSource.Length)
                        {
                            foreach (Run r in CreateSyntaxHighlightedRuns(
                                normalizedSource.Substring(startIdx + startLen)))
                                para.Inlines.Add(r);
                        }

                        startPositionText.Text = "pos: " + startIdx;
                        startPositionText.Foreground = Brushes.Green;
                    }
                    else
                    {
                        foreach (Run r in CreateSyntaxHighlightedRuns(normalizedSource))
                            para.Inlines.Add(r);

                        startPositionText.Text = "NOT FOUND";
                        startPositionText.Foreground = Brushes.Red;
                    }
                    stopPositionText.Text = "";
                }
                else
                {
                    foreach (Run r in CreateSyntaxHighlightedRuns(source))
                        para.Inlines.Add(r);

                    startPositionText.Text = "";
                    stopPositionText.Text = "";
                }
            }
            catch (Exception ex)
            {
                para.Inlines.Clear();
                try
                {
                    foreach (Run r in CreateSyntaxHighlightedRuns(source))
                        para.Inlines.Add(r);
                }
                catch (Exception)
                {
                    Run fallbackRun = new Run(source);
                    fallbackRun.Foreground = new SolidColorBrush(currentTheme.SourceTagColor);
                    para.Inlines.Add(fallbackRun);
                }

                startPositionText.Text = "Error: " + ex.Message;
                startPositionText.Foreground = Brushes.Red;
                stopPositionText.Text = "";
            }

            sourceView.Document.Blocks.Add(para);

            // Scroll to the start string position
            if (!string.IsNullOrEmpty(startStr))
            {
                try
                {
                    string normalizedForScroll = source.Replace("\r\n", "\n").Replace("\r", "\n");
                    int scrollIdx = -1;

                    if (TrackItem.IsRegexPattern(startStr))
                    {
                        var match = TrackItem.FindRegexMatch(normalizedForScroll, startStr, 0);
                        if (match != null && match.Success)
                            scrollIdx = match.Index;
                    }
                    else
                    {
                        string normalizedStartForScroll = startStr.Replace("\r\n", "\n").Replace("\r", "\n");
                        scrollIdx = normalizedForScroll.IndexOf(normalizedStartForScroll,
                            StringComparison.Ordinal);

                        if (scrollIdx < 0)
                        {
                            string cs = NormalizeWhitespace(normalizedForScroll);
                            string cst = NormalizeWhitespace(normalizedStartForScroll);
                            int ci = cs.IndexOf(cst, StringComparison.Ordinal);
                            if (ci >= 0)
                                scrollIdx = MapNormalizedToRaw(normalizedForScroll, ci);
                        }
                    }

                    if (scrollIdx >= 0)
                    {
                        ScrollSourceViewToPosition(scrollIdx);
                    }
                }
                catch (Exception) { }
            }
        }

        private void GoToStartString_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentSource) || string.IsNullOrEmpty(editStartString.Text))
            {
                statusFile.Text = "No source loaded or Start String is empty";
                return;
            }

            // Re-display source (this highlights and sets position texts)
            DisplaySource(currentSource);

            // Now scroll to the start position
            string normalizedSource = currentSource.Replace("\r\n", "\n").Replace("\r", "\n");
            string searchStr = editStartString.Text;
            int idx = -1;

            if (TrackItem.IsRegexPattern(searchStr))
            {
                var match = TrackItem.FindRegexMatch(normalizedSource, searchStr, 0);
                if (match != null && match.Success)
                    idx = match.Index;
            }
            else
            {
                string normalizedSearch = searchStr.Replace("\r\n", "\n").Replace("\r", "\n");
                idx = normalizedSource.IndexOf(normalizedSearch, StringComparison.Ordinal);

                if (idx < 0)
                {
                    string cs = NormalizeWhitespace(normalizedSource);
                    string cst = NormalizeWhitespace(normalizedSearch);
                    int ci = cs.IndexOf(cst, StringComparison.Ordinal);
                    if (ci >= 0)
                        idx = MapNormalizedToRaw(normalizedSource, ci);
                }
            }

            if (idx < 0)
            {
                statusFile.Text = "Start String not found in source";
                startPositionText.Text = "NOT FOUND";
                startPositionText.Foreground = Brushes.Red;
                return;
            }

            startPositionText.Text = "pos: " + idx;
            startPositionText.Foreground = Brushes.Green;
            ScrollSourceViewToPosition(idx);

            string regexNote = TrackItem.IsRegexPattern(searchStr) ? " (regex)" : "";
            statusFile.Text = "Start String found at position " + idx + regexNote;
        }

        private void GoToStopString_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(currentSource) || string.IsNullOrEmpty(editStopString.Text))
            {
                statusFile.Text = "No source loaded or Stop String is empty";
                return;
            }

            // Re-display source (this highlights and sets position texts)
            DisplaySource(currentSource);

            // Now scroll to the stop position
            string normalizedSource = currentSource.Replace("\r\n", "\n").Replace("\r", "\n");
            string startStr = editStartString.Text ?? "";
            string stopStr = editStopString.Text;
            int idx = -1;

            // Find start first so we search for stop after it
            int searchAfter = 0;
            if (!string.IsNullOrEmpty(startStr))
            {
                if (TrackItem.IsRegexPattern(startStr))
                {
                    var match = TrackItem.FindRegexMatch(normalizedSource, startStr, 0);
                    if (match != null && match.Success)
                        searchAfter = match.Index + match.Length;
                }
                else
                {
                    string ns = startStr.Replace("\r\n", "\n").Replace("\r", "\n");
                    int si = normalizedSource.IndexOf(ns, StringComparison.Ordinal);
                    if (si >= 0)
                        searchAfter = si + ns.Length;
                    else
                    {
                        string cs = NormalizeWhitespace(normalizedSource);
                        string cst = NormalizeWhitespace(ns);
                        int ci = cs.IndexOf(cst, StringComparison.Ordinal);
                        if (ci >= 0)
                        {
                            int rawStart = MapNormalizedToRaw(normalizedSource, ci);
                            int rawEnd = MapNormalizedToRaw(normalizedSource, ci + cst.Length);
                            searchAfter = rawEnd;
                        }
                    }
                }
            }

            if (TrackItem.IsRegexPattern(stopStr))
            {
                var match = TrackItem.FindRegexMatch(normalizedSource, stopStr, searchAfter);
                if (match != null && match.Success)
                    idx = match.Index;
            }
            else
            {
                string normalizedSearch = stopStr.Replace("\r\n", "\n").Replace("\r", "\n");
                idx = normalizedSource.IndexOf(normalizedSearch, searchAfter, StringComparison.Ordinal);

                if (idx < 0)
                {
                    string cs = NormalizeWhitespace(normalizedSource);
                    string cst = NormalizeWhitespace(normalizedSearch);
                    string collapsedBefore = NormalizeWhitespace(normalizedSource.Substring(0, searchAfter));
                    int collapsedSearchAfter = collapsedBefore.Length;

                    int ci = cs.IndexOf(cst, collapsedSearchAfter, StringComparison.Ordinal);
                    if (ci >= 0)
                        idx = MapNormalizedToRaw(normalizedSource, ci);
                }
            }

            if (idx < 0)
            {
                statusFile.Text = "Stop String not found in source";
                stopPositionText.Text = "NOT FOUND";
                stopPositionText.Foreground = Brushes.Red;
                return;
            }

            stopPositionText.Text = "pos: " + idx;
            stopPositionText.Foreground = Brushes.Green;
            ScrollSourceViewToPosition(idx);

            string regexNote = TrackItem.IsRegexPattern(stopStr) ? " (regex)" : "";
            statusFile.Text = "Stop String found at position " + idx + regexNote;
        }

		/// <summary>
		/// Scrolls the source view so that the text at the given character position
		/// in currentSource is visible, approximately 1/3 from the top of the view.
		/// Works by walking the RichTextBox runs and counting only text characters.
		/// </summary>
		private void ScrollSourceViewToPosition(int charPosition)
		{
		    try
		    {
		        TextPointer current = sourceView.Document.ContentStart;
		        int counted = 0;

		        while (current != null && current.CompareTo(sourceView.Document.ContentEnd) < 0)
		        {
		            if (current.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
		            {
		                int runLength = current.GetTextRunLength(LogicalDirection.Forward);

		                if (counted + runLength >= charPosition)
		                {
		                    int offsetInRun = charPosition - counted;
		                    TextPointer target = current.GetPositionAtOffset(offsetInRun);
		                    if (target != null)
		                    {
		                        Rect rect = target.GetCharacterRect(LogicalDirection.Forward);
		                        if (!rect.IsEmpty)
		                        {
		                            sourceView.ScrollToVerticalOffset(
		                                sourceView.VerticalOffset + rect.Top - sourceView.ActualHeight / 3);
		                        }
		                    }
		                    return;
		                }

		                counted += runLength;
		            }
		            current = current.GetNextContextPosition(LogicalDirection.Forward);
		        }
		    }
		    catch (Exception) { }
		}

		/// <summary>
		/// Collapses all runs of whitespace (spaces, tabs, \r, \n) into a single space.
		/// This makes string matching work regardless of how line endings or indentation differ.
		/// </summary>
		private static string NormalizeWhitespace(string input)
		{
		    if (string.IsNullOrEmpty(input))
		        return "";

		    StringBuilder sb = new StringBuilder(input.Length);
		    bool inWhitespace = false;

		    for (int i = 0; i < input.Length; i++)
		    {
		        char c = input[i];
		        if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
		        {
		            if (!inWhitespace)
		            {
		                sb.Append(' ');
		                inWhitespace = true;
		            }
		            // Skip additional whitespace chars
		        }
		        else
		        {
		            sb.Append(c);
		            inWhitespace = false;
		        }
		    }

		    return sb.ToString();
		}

		/// <summary>
		/// Maps a position in whitespace-normalized text back to the position in the raw text.
		/// </summary>
		private static int MapNormalizedToRaw(string raw, int normalizedPos)
		{
		    int normIdx = 0;
		    bool inWhitespace = false;

		    for (int rawI = 0; rawI < raw.Length; rawI++)
		    {
		        if (normIdx >= normalizedPos)
		            return rawI;

		        char c = raw[rawI];
		        if (c == ' ' || c == '\t' || c == '\r' || c == '\n')
		        {
		            if (!inWhitespace)
		            {
		                normIdx++; // The single space we collapsed to
		                inWhitespace = true;
		            }
		        }
		        else
		        {
		            normIdx++;
		            inWhitespace = false;
		        }
		    }

		    return raw.Length;
		}

        private void ScrollToPosition(int charPosition, int length)
        {
            try
            {
                TextPointer start = sourceView.Document.ContentStart;
                TextPointer current = start;
                int count = 0;

                // Walk through the document to find the right position
                while (current != null && count < charPosition)
                {
                    if (current.GetPointerContext(LogicalDirection.Forward) == TextPointerContext.Text)
                    {
                        string text = current.GetTextInRun(LogicalDirection.Forward);
                        int remaining = charPosition - count;

                        if (remaining <= text.Length)
                        {
                            TextPointer target = current.GetPositionAtOffset(remaining);
                            if (target != null)
                            {
                                sourceView.CaretPosition = target;
                                Rect rect = target.GetCharacterRect(LogicalDirection.Forward);
                                sourceView.ScrollToVerticalOffset(
                                    sourceView.VerticalOffset + rect.Top - sourceView.ActualHeight / 3);
                            }
                            return;
                        }

                        count += text.Length;
                    }
                    current = current.GetNextContextPosition(LogicalDirection.Forward);
                }
            }
            catch (Exception)
            {
            }
        }

        private async void InitWebView()
        {
            try
            {
                rightTabs.SelectedIndex = 1;
                await System.Threading.Tasks.Task.Delay(100);

                var options = new Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions();
                options.AreBrowserExtensionsEnabled = true;

                var environment = await Microsoft.Web.WebView2.Core.CoreWebView2Environment.CreateAsync(
                    null, webViewPath, options);

                await webView.EnsureCoreWebView2Async(environment);
                // Enables password save confirmation
                webView.CoreWebView2.Profile.IsPasswordAutosaveEnabled = true;

                // Inject cookie popup blocker after each page loads
                webView.CoreWebView2.NavigationCompleted += async (s2, e2) =>
                {
                    if (blockCookiePopups)
                    {
                        await InjectCookiePopupBlocker();
                    }
                };

                // Switch back to source tab and load the first item's source
                rightTabs.SelectedIndex = 0;
                await System.Threading.Tasks.Task.Delay(200);

                if (currentTrackItem != null && !string.IsNullOrWhiteSpace(currentTrackItem.TrackURL))
                {
                    AutoDownloadSource(currentTrackItem.TrackURL);
                }
            }
            catch (Exception)
            {
            }
        }

        private async System.Threading.Tasks.Task InjectCookiePopupBlocker()
        {
            try
            {
                string script = @"
                    (function() {
                        if (window.__patCookieBlockerActive) return;
                        window.__patCookieBlockerActive = true;

                        var selectors = [
                            '[id*=""cookie"" i]', '[class*=""cookie"" i]',
                            '[id*=""consent"" i]', '[class*=""consent"" i]',
                            '[id*=""gdpr"" i]', '[class*=""gdpr"" i]',
                            '[class*=""cc-banner"" i]', '[class*=""cc-window"" i]',
                            '[id*=""onetrust"" i]', '[class*=""onetrust"" i]',
                            '[id*=""CybotCookiebot"" i]',
                            '[class*=""cookiewall"" i]', '[id*=""cookiewall"" i]',
                            '.cookie-banner', '.cookie-notice', '.cookie-popup',
                            '#cookie-law-info-bar', '.cli-modal',
                            '[aria-label*=""cookie"" i]', '[aria-label*=""consent"" i]'
                        ];

                        var rejectKeywords = ['reject', 'decline', 'deny', 'refuse', 'necessary only', 'essential only'];
                        var acceptKeywords = ['accept all', 'accept', 'agree', 'got it', 'i understand'];

                        function killPopups() {
                            selectors.forEach(function(sel) {
                                try {
                                    document.querySelectorAll(sel).forEach(function(el) {
                                        if (el.offsetHeight > 0) {
                                            el.style.setProperty('display', 'none', 'important');
                                            el.style.setProperty('visibility', 'hidden', 'important');
                                            el.style.setProperty('opacity', '0', 'important');
                                            el.style.setProperty('pointer-events', 'none', 'important');
                                        }
                                    });
                                } catch(e) {}
                            });

                            document.body.style.setProperty('overflow', 'auto', 'important');
                            document.documentElement.style.setProperty('overflow', 'auto', 'important');

                            var allButtons = document.querySelectorAll('button, a[role=""button""], input[type=""button""], input[type=""submit""]');
                            var clicked = false;

                            for (var i = 0; i < allButtons.length && !clicked; i++) {
                                var text = (allButtons[i].innerText || allButtons[i].value || '').toLowerCase().trim();
                                for (var j = 0; j < rejectKeywords.length; j++) {
                                    if (text.includes(rejectKeywords[j])) {
                                        allButtons[i].click();
                                        clicked = true;
                                        break;
                                    }
                                }
                            }
                        }

                        // Run immediately
                        killPopups();

                        // Run on short delays for lazy-loaded popups
                        setTimeout(killPopups, 500);
                        setTimeout(killPopups, 1500);
                        setTimeout(killPopups, 3000);
                        setTimeout(killPopups, 5000);

                        // Watch for new elements being added to the DOM
                        var observer = new MutationObserver(function(mutations) {
                            killPopups();
                        });
                        observer.observe(document.body, { childList: true, subtree: true });

                        // Also inject CSS to hide common patterns immediately
                        var style = document.createElement('style');
                        style.id = '__patCookieBlockerCSS';
                        style.textContent = `
                            [id*='cookie' i], [class*='cookie' i],
                            [id*='consent' i], [class*='consent' i],
                            [id*='gdpr' i], [class*='gdpr' i],
                            [id*='onetrust' i], [class*='onetrust' i],
                            [id*='CybotCookiebot' i],
                            .cookie-banner, .cookie-notice, .cookie-popup,
                            #cookie-law-info-bar, .cli-modal,
                            [class*='cc-banner' i], [class*='cc-window' i],
                            [class*='cookiewall' i], [id*='cookiewall' i] {
                                display: none !important;
                                visibility: hidden !important;
                                opacity: 0 !important;
                                pointer-events: none !important;
                            }
                            body, html {
                                overflow: auto !important;
                            }
                        `;
                        document.head.appendChild(style);
                    })();
                ";

                await webView.CoreWebView2.ExecuteScriptAsync(script);
            }
            catch (Exception)
            {
            }
        }

        private async void ClearCookies_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult confirm = MessageBox.Show(
                "Delete all cookies?\n\n" +
                "This will sign you out of all websites in the WebView.",
                "Clear Cookies", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes)
                return;

            try
            {
                await webView.EnsureCoreWebView2Async();
                webView.CoreWebView2.CookieManager.DeleteAllCookies();

                MessageBox.Show("All cookies have been deleted.\n\n" +
                    "You may need to reload any open pages.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
                statusFile.Text = "All cookies cleared";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error clearing cookies:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void ClearAllData_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult confirm = MessageBox.Show(
                "Clear ALL browsing data?\n\n" +
                "This will delete:\n" +
                "• All cookies\n" +
                "• Browser cache\n" +
                "• Local storage & session storage\n" +
                "• All cached files\n\n" +
                "You will be signed out of all websites.",
                "Clear All Data", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
                return;

            try
            {
                await webView.EnsureCoreWebView2Async();

                // Delete cookies
                webView.CoreWebView2.CookieManager.DeleteAllCookies();

                // Clear cache and data via DevTools protocol
                await webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Network.clearBrowserCache", "{}");
                await webView.CoreWebView2.CallDevToolsProtocolMethodAsync(
                    "Network.clearBrowserCookies", "{}");

                // Clear storage via JavaScript
                await webView.CoreWebView2.ExecuteScriptAsync(@"
                    try { localStorage.clear(); } catch(e) {}
                    try { sessionStorage.clear(); } catch(e) {}
                    try {
                        if (window.caches) {
                            caches.keys().then(function(names) {
                                names.forEach(function(name) { caches.delete(name); });
                            });
                        }
                    } catch(e) {}
                ");

                MessageBox.Show("All browsing data has been cleared.\n\n" +
                    "You may need to reload any open pages.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
                statusFile.Text = "All browsing data and cookies cleared";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error clearing data:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // ============================
        // Column Settings
        // ============================

        private void SaveColumnSettings()
        {
            GridView gv = itemList.View as GridView;
            if (gv == null) return;

            // Get the full list of columns (visible + hidden) from defaults
            List<ColumnSetting> allDefaults = WindowSettings.GetDefaultColumns();
            List<ColumnSetting> saved = new List<ColumnSetting>();

            // First, add visible columns in their current order with current widths
            int displayIndex = 0;
            foreach (GridViewColumn col in gv.Columns)
            {
                // Skip checkbox and status columns (first two)
                string header = col.Header as string;
                if (header == null) continue;
                if (header == "St") continue;

                // Find the binding for this column
                string binding = "";
                foreach (ColumnSetting def in allDefaults)
                {
                    if (def.Header == header)
                    {
                        binding = def.Binding;
                        break;
                    }
                }

                if (!string.IsNullOrEmpty(binding))
                {
                    saved.Add(new ColumnSetting
                    {
                        Header = header,
                        Binding = binding,
                        Width = col.Width > 0 ? col.Width : col.ActualWidth,
                        Visible = true,
                        DisplayIndex = displayIndex++
                    });
                }
            }

            // Now add hidden columns (those in defaults but not in the gridview)
            foreach (ColumnSetting def in allDefaults)
            {
                bool found = false;
                foreach (ColumnSetting s in saved)
                {
                    if (s.Binding == def.Binding)
                    {
                        found = true;
                        break;
                    }
                }

                if (!found)
                {
                    saved.Add(new ColumnSetting
                    {
                        Header = def.Header,
                        Binding = def.Binding,
                        Width = def.Width,
                        Visible = false,
                        DisplayIndex = displayIndex++
                    });
                }
            }

            windowSettings.ColumnSettings = saved;
        }

        private void ApplyColumnSettings_Click(object sender, RoutedEventArgs e)
        {
            // Read checkbox states and update settings
            List<ColumnSetting> allDefaults = WindowSettings.GetDefaultColumns();

            // First save current column widths/order for visible columns
            SaveColumnSettings();

            // Now apply checkbox visibility changes
            List<ColumnSetting> updated = new List<ColumnSetting>();
            int displayIndex = 0;

            foreach (var child in columnCheckboxPanel.Children)
            {
                CheckBox chk = child as CheckBox;
                if (chk == null) continue;

                string binding = chk.Tag as string;
                if (string.IsNullOrEmpty(binding)) continue;

                // Find existing setting or default
                ColumnSetting existing = null;
                foreach (ColumnSetting cs in windowSettings.ColumnSettings)
                {
                    if (cs.Binding == binding)
                    {
                        existing = cs;
                        break;
                    }
                }

                if (existing == null)
                {
                    foreach (ColumnSetting def in allDefaults)
                    {
                        if (def.Binding == binding)
                        {
                            existing = def;
                            break;
                        }
                    }
                }

                if (existing != null)
                {
                    updated.Add(new ColumnSetting
                    {
                        Header = existing.Header,
                        Binding = existing.Binding,
                        Width = existing.Width,
                        Visible = chk.IsChecked == true,
                        DisplayIndex = displayIndex++
                    });
                }
            }

            windowSettings.ColumnSettings = updated;
            RebuildColumns();
            statusFile.Text = "Column settings applied";
        }

        private void SaveAppSettings_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SaveColumnSettings();
                SaveWindowSettings();
                statusFile.Text = "All settings saved";
                MessageBox.Show("Settings saved successfully.",
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving settings:\n" + ex.Message,
                    "PAT v7", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RestoreDefaultColumns_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult confirm = MessageBox.Show(
                "Restore all column settings to defaults?\n\n" +
                "This will reset visibility, widths, and order.",
                "Restore Defaults", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes)
                return;

            windowSettings.ColumnSettings = WindowSettings.GetDefaultColumns();

            // Update checkboxes
            columnCheckboxPanel.Children.Clear();
            List<ColumnSetting> defaults = WindowSettings.GetDefaultColumns();
            foreach (ColumnSetting cs in defaults)
            {
                CheckBox chk = new CheckBox();
                chk.Content = cs.Header + "  (" + cs.Binding + ")";
                chk.IsChecked = cs.Visible;
                chk.Tag = cs.Binding;
                chk.Margin = new Thickness(0, 2, 0, 2);
                columnCheckboxPanel.Children.Add(chk);
            }

            RebuildColumns();
            statusFile.Text = "Column settings restored to defaults";
        }

        // ============================
        // Source Tab - Search
        // ============================
        /// <summary>
        /// Parses an HTML string and returns a list of Runs with simple syntax highlighting.
        /// Everything inside angle brackets is one color, everything outside is another.
        /// </summary>
		private List<Run> CreateSyntaxHighlightedRuns(string html)
		{
		    List<Run> runs = new List<Run>();
		    if (string.IsNullOrEmpty(html))
		        return runs;

		    SolidColorBrush tagBrush = new SolidColorBrush(currentTheme.SourceTagColor);
		    SolidColorBrush textBrush = new SolidColorBrush(currentTheme.SourceTextColor);

            int i = 0;

            while (i < html.Length)
            {
                if (html[i] == '<')
                {
                    // Find the closing >
                    int end = html.IndexOf('>', i);
                    if (end < 0) end = html.Length - 1;

                    runs.Add(MakeRun(html.Substring(i, end - i + 1), tagBrush));
                    i = end + 1;
                }
                else
                {
                    // Find the next <
                    int next = html.IndexOf('<', i);
                    if (next < 0) next = html.Length;

                    runs.Add(MakeRun(html.Substring(i, next - i), textBrush));
                    i = next;
                }
            }

            return runs;
        }

        private Run MakeRun(string text, SolidColorBrush brush)
        {
            Run run = new Run(text);
            run.Foreground = brush;
            return run;
        }

//         private void SearchTop_Click(object sender, RoutedEventArgs e)
//         {
//             string searchText = findString.Text;
//             if (string.IsNullOrEmpty(searchText) || string.IsNullOrEmpty(currentSource))
//                 return;

//             searchHitPositions.Clear();
//             int pos = 0;
//             string lowerSource = currentSource.ToLower();
//             string lowerSearch = searchText.ToLower();

//             while ((pos = lowerSource.IndexOf(lowerSearch, pos)) >= 0)
//             {
//                 searchHitPositions.Add(pos);
//                 pos += lowerSearch.Length;
//             }

//             findCount.Text = searchHitPositions.Count + " found";

//             if (searchHitPositions.Count > 0)
//             {
//                 currentSearchIndex = 0;
//                 HighlightSearch(searchText);
//             }
//         }

		private void SearchDown_Click(object sender, RoutedEventArgs e)
		{
		    if (string.IsNullOrEmpty(findString.Text) || string.IsNullOrEmpty(currentSource))
		        return;

		    if (searchHitPositions.Count == 0)
		    {
		        searchHitPositions.Clear();
		        int pos = 0;
		        string lowerSource = currentSource.ToLower();
		        string lowerSearch = findString.Text.ToLower();

		        while ((pos = lowerSource.IndexOf(lowerSearch, pos)) >= 0)
		        {
		            searchHitPositions.Add(pos);
		            pos += lowerSearch.Length;
		        }

		        findCount.Text = searchHitPositions.Count + " found";

		        if (searchHitPositions.Count > 0)
		        {
		            currentSearchIndex = 0;
		            HighlightSearch(findString.Text);
		        }
		        return;
		    }

		    currentSearchIndex++;
		    if (currentSearchIndex >= searchHitPositions.Count)
		        currentSearchIndex = 0;

		    HighlightSearch(findString.Text);
		}

		private void SearchUp_Click(object sender, RoutedEventArgs e)
		{
		    if (string.IsNullOrEmpty(findString.Text) || string.IsNullOrEmpty(currentSource))
		        return;

		    if (searchHitPositions.Count == 0)
		    {
		        searchHitPositions.Clear();
		        int pos = 0;
		        string lowerSource = currentSource.ToLower();
		        string lowerSearch = findString.Text.ToLower();

		        while ((pos = lowerSource.IndexOf(lowerSearch, pos)) >= 0)
		        {
		            searchHitPositions.Add(pos);
		            pos += lowerSearch.Length;
		        }

		        findCount.Text = searchHitPositions.Count + " found";

		        if (searchHitPositions.Count > 0)
		        {
		            currentSearchIndex = searchHitPositions.Count - 1;
		            HighlightSearch(findString.Text);
		        }
		        return;
		    }

		    currentSearchIndex--;
		    if (currentSearchIndex < 0)
		        currentSearchIndex = searchHitPositions.Count - 1;

		    HighlightSearch(findString.Text);
		}

        private void HighlightSearch(string searchText)
        {
            if (searchHitPositions.Count == 0 || currentSearchIndex < 0)
                return;

            sourceView.Document.Blocks.Clear();

            Paragraph para = new Paragraph();
            para.FontFamily = new FontFamily("Consolas");
            para.FontSize = sourceView.FontSize;
            para.Margin = new Thickness(0);

            int lastPos = 0;
            string lowerSource = currentSource.ToLower();
            string lowerSearch = searchText.ToLower();
            int hitIndex = 0;
            int pos = 0;

            while ((pos = lowerSource.IndexOf(lowerSearch, lastPos)) >= 0)
            {
                if (pos > lastPos)
                {
                    para.Inlines.Add(new Run(currentSource.Substring(lastPos, pos - lastPos)));
                }

                Run hitRun = new Run(currentSource.Substring(pos, searchText.Length));

                if (hitIndex == currentSearchIndex)
                {
                    hitRun.Background = Brushes.Orange;
                    hitRun.Foreground = Brushes.Black;
                    hitRun.FontWeight = FontWeights.Bold;
                }
                else
                {
                    hitRun.Background = Brushes.LightYellow;
                    hitRun.Foreground = Brushes.Black;
                }

                para.Inlines.Add(hitRun);

                lastPos = pos + searchText.Length;
                hitIndex++;
            }

            if (lastPos < currentSource.Length)
            {
                para.Inlines.Add(new Run(currentSource.Substring(lastPos)));
            }

            sourceView.Document.Blocks.Add(para);

            findCount.Text = (currentSearchIndex + 1) + " / " + searchHitPositions.Count;

            try
            {
                TextPointer start = sourceView.Document.ContentStart;
                TextPointer pointer = start.GetPositionAtOffset(
                    searchHitPositions[currentSearchIndex] + hitIndex * 2);
                if (pointer != null)
                {
                    sourceView.CaretPosition = pointer;
                    Rect rect = pointer.GetCharacterRect(LogicalDirection.Forward);
                    sourceView.ScrollToVerticalOffset(
                        sourceView.VerticalOffset + rect.Top - sourceView.ActualHeight / 3);
                }
            }
            catch (Exception)
            {
            }
        }

        // ============================
        // WebView Tab
        // ============================

        private bool webViewAutoZoom = true;
        private double webViewBaseWidth = 1024.0;
        private bool webViewReady = false;

        private void SetupWebViewAutoZoom()
        {
            webView.CoreWebView2InitializationCompleted += (s, e) =>
            {
                if (e.IsSuccess)
                {
                    webViewReady = true;

                    webView.CoreWebView2.NavigationCompleted += (s2, e2) =>
                    {
                        if (webViewAutoZoom)
                        {
                            ApplyAutoZoom();
                        }
                    };
                }
            };

            webView.SizeChanged += (s, e) =>
            {
                if (webViewAutoZoom && webViewReady && e.NewSize.Width > 0)
                {
                    ApplyAutoZoom();
                }
            };
        }

        private void ApplyAutoZoom()
        {
            if (webView.ActualWidth > 0 && webView.CoreWebView2 != null)
            {
                double zoom = webView.ActualWidth / webViewBaseWidth;
                zoom = Math.Max(0.25, Math.Min(2.0, zoom));
                webView.ZoomFactor = zoom;
            }
        }

        private void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            webViewAutoZoom = false;
            if (webView.CoreWebView2 != null)
                webView.ZoomFactor += 0.1;
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            webViewAutoZoom = false;
            if (webView.CoreWebView2 != null)
                webView.ZoomFactor = Math.Max(0.1, webView.ZoomFactor - 0.1);
        }

        private void ZoomReset_Click(object sender, RoutedEventArgs e)
        {
            webViewAutoZoom = false;
            if (webView.CoreWebView2 != null)
                webView.ZoomFactor = 1.0;
        }

        private void ZoomFit_Click(object sender, RoutedEventArgs e)
        {
            webViewAutoZoom = true;
            ApplyAutoZoom();
        }
    }
}