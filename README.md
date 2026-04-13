# SPS Published App Track (PAT) x64

- **v7.x.x.x** — Complete major rewrite by **sl23**
- **v6.x.x.x** — Major x64 rewrite by **sl23**
- **v5.2.0.2** — Original code by **VVV_Easy_SyMenu**

---

## About

Originally designed solely to be used with SyMenu's[SyMenu](https://www.ugmfree.it/) and its SPS Builder app, it was a basic app and a quick fix for a unique problem: Checking many websites for app updates to manage portable application suites.

After an update to bring it up to date with 64bit support and enhancing the looks and broadening functionality, I then decided to merge my generic Rainmeter updater skin functionality with SPS PAT.

Now it has become a fairly unique app that can be used on just about any website, to monitor changes in text. You simply create a "Track File" and enter the site, start/stop strings and the app can then search those sites for changes between the two strings. The uses are limited only by you! 

In theory, this should work on any site as it has two modes: HTML scraping and Rendered text.
HTML scraping allows more complex strings to be used, including RegEx terms, to search for required data, if that data changes onsite, then a scan will show you what's changed.

This app also has built in WebView2 functionality that allows for simplistic browsing, passwords are saved, cookie control, ublock origin support. I'm hoping to add more extensions to it, for a more custom experience, as well as more functionality.

Some examples:
- Email site updates
- Checking sites for app/vst versions
- Check stock sites for changes
- Check feeds for updates
- etc....

---

## User-Agent Strings

If any of your tracked URLs start failing because the server rejects your User-Agent as outdated or suspicious, you can grab a current User-Agent string from [useragents.io](https://useragents.io/explore) and update the one in the code.

---

## Changelog

### v7.0.4.480 — 2026.04.09
- Toolbar overflow button restored
- Added WebView toolbar button to add current page to Track URL field
- Added new check function taht search for plain text in Source tab
- Added WebView context for adding Download URL/Start/Stop strings direct to the fields in Track Settings.
   Download URL is context sensitive, ie, works on file links only
- Window title now includes auto-version info
- Save buttons disabled by default, only enabled when Track Settings changes are detected
- Disabled unused menutitems to alert users they aren't in use yet
- Some other minor enhancements:


### v7.0.0.0 — 2026.04.07 (sl23)

Major upgrade — now compiled as **x64**.

- Completely rewritten in C# WPF using .NET SDK v8.0.25
- Created separate panels for each section
- Icons sue Segoe Fluent Font Icons
- Full theme support added
- App tracks are now stored individually in their own .track files
- ListView column moving, sorting, sizing, and hiding
- Simple syntax highlighting for Source Tab
- WebView2 added to allow direct page viewing in the app
- Added Ublock Origin support in WebView2 
- Categories can be added to allow better organisation of tracked apps
- Added browsing functions to WebView

### Todo ###
- Add full SyMenu SPS support
- Sort out MenuBar items not working
- Code cleanup
- Possible performance enhancements

---

### v6.0.0.0 — 2026.03.30 (sl23)

Major upgrade — now compiled as **x64**.

- Merged windows into a single split window
- Updated text and buttons
- Added icons to buttons
- Rearranged layout
- Added new columns
- Added ListView column moving, sorting, sizing, and hiding
- Improved search functions to search all suites
- Can now press Enter to search
- Publisher search now uses fuzzy matching
- Added Dark Mode
- Resized defaults
- Changed settings extension to XML
- Improved performance and GUI responsiveness
- Updated Help file
- PAT now works from any location, not just the default SyMenuSuite location
- Open file now uses modern Explorer windows
- Toggle panel — open/close
- Toggle panel between right and bottom orientation
- Added button to toggle between RTF and HTML view
- Edited SPS can be saved back to the your suites zip files
- Extraction enhanced, more secure so you dont lose modified SPS in tmp folder

---

### Original code and changelog by VVV_Easy_SyMenu

#### v5.2.0.2 — 2025.03.12
- Merged Form1 and Form2 into single window with SplitContainer

#### v5.2.0.1 — 2025.01.14
- Updated user agent strings

#### v5.2 — 2018.03.01
- Corrected bugs when changing colours in second track or changing strings in edit track

#### v5.1 — 2018.02.27
- Download size test improvements (now works with SourceForge)

#### v5.0 — 2018.02.25
- TLS 1.2 protocol supported (requires .NET Framework v4.6.1)
- Download size test improvements

#### v4.0 — 2017.02.26
- Added menu bar and config file `ConfigPAT.xml`
- Help now opens forum topic

#### v3.0 — 2017.02.05
- Using contextual menu
- Allow several Edit forms

#### v2.0 — 2017.02.03
- Now only SPS App flavour (no Launcher needed), named SPS Published App Track (PAT)
- Added SPS Builder call with local SPS file copy (temporarily located in `SyMenuSuite\_Trash\_TmpPAT`)

#### v1.4b — 2017.01.11
- Corrected bug saving files with the plugin executed with Launcher (in the SPS app flavour)

#### v1.4 — 2017.01.09
- Showed version in the window title
- Added more search options
- Corrected some bugs

#### v1.3 — 2017.01.05
- Added Tooltips
- Manage SPS or ZIP `_Cache` SPS Suite files (SyMenu version > 5.07.6190 [2016.12.13])
- Added SPS Publisher column (`<SPSPublisherName>` becomes `<SPSTrackerName>`)
- Allows several SPS Publisher names in the SPS Tracker Name

#### v1.2 — 2016.12.18
- Now available as SPS standalone program: SyMenu Published App Track (Others — Specialized Editors). Thanks Gian.

#### v1.2 — 2016.11.10
- Corrected some bugs
- Full automatic SyMenu plugin detection

#### v1.1 — 2016.10.10
- Added App Icon, Version, and Release Date
- Corrected some bugs
- Known issue: No automatic SyMenu plugin detection

#### v0.1 — 2016.10.02
- First published version

#### Beta — 2016.09.21
- Initial beta release

---

## Credits

- **VVV_Easy_SyMenu** — Original code and all versions up to v5.2.0.2
- **sl23** — v6.x.x.x and above
