# Elemental Tracker (ET)
Precision web tracking made simple to keep you informed of what matters to you.

---
- **v7.x.x.x** — Complete major rewrite by **sl23**
- **v6.x.x.x** — Major x64 rewrite by **sl23**
- **v5.2.0.2** — Original code by **VVV_Easy_SyMenu**
---

## About

Originally designed solely to be used with [SyMenu](https://www.ugmfree.it/) and its SPS Builder app, it was a basic app and a quick fix for a unique problem: Checking many websites for app updates to manage portable application suites.

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

## Credits

- **VVV_Easy_SyMenu** — Original code and all versions up to v5.2.0.2
- **sl23** — v6.x.x.x and above
