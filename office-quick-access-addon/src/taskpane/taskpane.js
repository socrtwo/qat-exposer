/*
 * Office Quick Access Add-in - Task Pane
 * Single dropdown exposing every QAT-customizable command.
 * Commands that Office.js can execute run directly;
 * the rest show the ribbon path or keyboard shortcut.
 */

/* global Office, Word, Excel, PowerPoint */

(function () {
  "use strict";

  // ── Toast ──────────────────────────────────────────────────────────
  var toastEl, toastTimer;
  function showToast(msg) {
    if (!toastEl) {
      toastEl = document.createElement("div");
      toastEl.className = "toast";
      document.body.appendChild(toastEl);
    }
    toastEl.textContent = msg;
    toastEl.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove("show"); }, 2500);
  }

  // ── Ribbon / shortcut hint helper ──────────────────────────────────
  function ribbon(location) { return function () { showToast("Ribbon: " + location); }; }
  function shortcut(keys) { return function () { showToast("Shortcut: " + keys); }; }
  function notAvailable(reason) { return function () { showToast(reason || "Not available via add-in API."); }; }

  // ── Office.js helpers ──────────────────────────────────────────────
  function isWord()  { return Office.context.host === Office.HostType.Word; }
  function isExcel() { return Office.context.host === Office.HostType.Excel; }

  function wordRun(fn, ok, label) {
    Word.run(fn).then(function () { showToast(ok || (label + " done.")); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }
  function excelRun(fn, ok, label) {
    Excel.run(fn).then(function () { showToast(ok || (label + " done.")); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }

  function toggleWordFont(prop) {
    wordRun(function (ctx) {
      var sel = ctx.document.getSelection(); sel.font.load(prop);
      return ctx.sync().then(function () { sel.font[prop] = !sel.font[prop]; return ctx.sync(); });
    }, prop.charAt(0).toUpperCase() + prop.slice(1) + " toggled.");
  }
  function toggleExcelFont(prop) {
    excelRun(function (ctx) {
      var r = ctx.workbook.getSelectedRange(); r.format.font.load(prop);
      return ctx.sync().then(function () { r.format.font[prop] = !r.format.font[prop]; return ctx.sync(); });
    }, prop.charAt(0).toUpperCase() + prop.slice(1) + " toggled.");
  }
  function toggleFont(prop) {
    if (isWord()) toggleWordFont(prop);
    else if (isExcel()) toggleExcelFont(prop);
    else showToast("Not available for this app.");
  }

  function setFontSize(sz) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().font.size = sz; return ctx.sync(); }, "Font size: " + sz);
    else if (isExcel()) excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.font.size = sz; return ctx.sync(); }, "Font size: " + sz);
    else showToast("Not available.");
  }
  function setFontColor(c) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().font.color = c; return ctx.sync(); }, "Color applied.");
    else if (isExcel()) excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.font.color = c; return ctx.sync(); }, "Color applied.");
    else showToast("Not available.");
  }
  function setFontName(name) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().font.name = name; return ctx.sync(); }, "Font: " + name);
    else if (isExcel()) excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.font.name = name; return ctx.sync(); }, "Font: " + name);
    else showToast("Not available.");
  }
  function setAlignment(a) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x.alignment = a; }); return ctx.sync(); });
    }, "Alignment: " + a);
    else if (isExcel()) excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.horizontalAlignment = a; return ctx.sync(); }, "Alignment: " + a);
    else showToast("Not available.");
  }
  function setHighlight(color) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().font.highlightColor = color; return ctx.sync(); }, color ? "Highlighted." : "Highlight removed.");
    else if (isExcel()) {
      if (color) excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.fill.color = color; return ctx.sync(); }, "Highlighted.");
      else excelRun(function (ctx) { ctx.workbook.getSelectedRange().format.fill.clear(); return ctx.sync(); }, "Highlight removed.");
    }
  }
  function setLineSpacing(val) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x.lineSpacing = val; }); return ctx.sync(); });
    }, "Line spacing: " + val);
    else showToast("Ribbon: Home > Line Spacing");
  }
  function setSpaceBefore(val) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x.spaceBefore = val; }); return ctx.sync(); });
    }, "Space before: " + val + "pt");
    else showToast("Not available.");
  }
  function setSpaceAfter(val) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x.spaceAfter = val; }); return ctx.sync(); });
    }, "Space after: " + val + "pt");
    else showToast("Not available.");
  }
  function insertBreak(type) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().insertBreak(type, "After"); return ctx.sync(); }, type + " break inserted.");
    else showToast("Not available for this app.");
  }
  function insertHtml(html, msg) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().insertHtml(html, "After"); return ctx.sync(); }, msg);
    else showToast("Not available for this app.");
  }
  function clearFormatting() {
    if (isWord()) wordRun(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.font.bold = false; sel.font.italic = false; sel.font.underline = "None";
      sel.font.strikethrough = false; sel.font.superscript = false; sel.font.subscript = false;
      sel.font.size = 11; sel.font.name = "Calibri"; sel.font.color = "#000000";
      sel.font.highlightColor = null;
      return ctx.sync();
    }, "Formatting cleared.");
    else if (isExcel()) excelRun(function (ctx) {
      var r = ctx.workbook.getSelectedRange();
      r.format.font.bold = false; r.format.font.italic = false; r.format.font.underline = "None";
      r.format.font.strikethrough = false; r.format.font.size = 11;
      r.format.font.name = "Calibri"; r.format.font.color = "#000000";
      return ctx.sync();
    }, "Formatting cleared.");
  }
  function wordInsertText(text, msg) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().insertText(text, "Replace"); return ctx.sync(); }, msg);
    else showToast("Not available.");
  }
  function setIndent(prop, val) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x[prop] = val; }); return ctx.sync(); });
    }, prop + " set.");
    else showToast("Not available.");
  }
  function changeIndent(delta) {
    if (isWord()) wordRun(function (ctx) {
      var p = ctx.document.getSelection().paragraphs; p.load("items/leftIndent");
      return ctx.sync().then(function () { p.items.forEach(function (x) { x.leftIndent = Math.max(0, x.leftIndent + delta); }); return ctx.sync(); });
    }, "Indent changed.");
    else showToast("Ribbon: Home > Increase/Decrease Indent");
  }
  function setStyle(name) {
    if (isWord()) wordRun(function (ctx) { ctx.document.getSelection().style = name; return ctx.sync(); }, "Style: " + name);
    else showToast("Not available.");
  }
  function wordSearch(text) {
    if (isWord()) wordRun(function (ctx) {
      var body = ctx.document.body;
      var results = body.search(text, { matchCase: false, matchWholeWord: false });
      results.load("items");
      return ctx.sync().then(function () {
        if (results.items.length > 0) { results.items[0].select(); return ctx.sync(); }
        else { showToast("Not found."); }
      });
    }, "Found.");
    else showToast("Use Ctrl+F.");
  }

  // ── The full QAT command list ──────────────────────────────────────
  // Each entry: [value, label, handler]
  // handler is a function, or null => falls back to ribbon hint embedded in label
  var ALL_COMMANDS = [
    // --- A ---
    ["acceptAllChanges", "Accept All Changes", ribbon("Review > Accept > Accept All Changes")],
    ["acceptChange", "Accept Change", ribbon("Review > Accept")],
    ["accessibility", "Accessibility Checker", ribbon("Review > Check Accessibility")],
    ["addChartElement", "Add Chart Element", ribbon("Chart Design > Add Chart Element")],
    ["addHorizontalLine", "Add Horizontal Line", function () { insertHtml('<hr style="border:1px solid #999;width:100%">', "Horizontal line inserted."); }],
    ["addShape", "Add Shape", ribbon("Insert > Shapes")],
    ["addText", "Add Text", ribbon("References > Add Text")],
    ["adjustListIndents", "Adjust List Indents", ribbon("Right-click list > Adjust List Indents")],
    ["advancedFind", "Advanced Find", shortcut("Ctrl+H")],
    ["alignBottom", "Align Bottom", ribbon("Table Layout > Cell Alignment > Align Bottom")],
    ["alignBottomCenter", "Align Bottom Center", ribbon("Table Layout > Cell Alignment")],
    ["alignBottomLeft", "Align Bottom Left", ribbon("Table Layout > Cell Alignment")],
    ["alignBottomRight", "Align Bottom Right", ribbon("Table Layout > Cell Alignment")],
    ["alignCenter", "Align Center", function () { setAlignment("Center"); }],
    ["alignJustify", "Align Justify", function () { setAlignment("Justified"); }],
    ["alignLeft", "Align Left", function () { setAlignment("Left"); }],
    ["alignMiddleCenter", "Align Middle Center", ribbon("Table Layout > Cell Alignment")],
    ["alignMiddleLeft", "Align Middle Left", ribbon("Table Layout > Cell Alignment")],
    ["alignMiddleRight", "Align Middle Right", ribbon("Table Layout > Cell Alignment")],
    ["alignRight", "Align Right", function () { setAlignment("Right"); }],
    ["alignTopCenter", "Align Top Center", ribbon("Table Layout > Cell Alignment")],
    ["alignTopLeft", "Align Top Left", ribbon("Table Layout > Cell Alignment")],
    ["alignTopRight", "Align Top Right", ribbon("Table Layout > Cell Alignment")],
    ["allCaps", "All Caps", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("allCaps");
        return ctx.sync().then(function () { s.font.allCaps = !s.font.allCaps; return ctx.sync(); });
      }, "All caps toggled.");
      else showToast("Not available.");
    }],
    ["arrangeAll", "Arrange All", ribbon("View > Arrange All")],
    ["attachTemplate", "Attach Template", ribbon("Developer > Document Template")],
    ["autoCorrectOptions", "AutoCorrect Options", ribbon("File > Options > Proofing > AutoCorrect Options")],
    ["autoFitContents", "AutoFit Contents", ribbon("Table Layout > AutoFit > AutoFit Contents")],
    ["autoFitWindow", "AutoFit Window", ribbon("Table Layout > AutoFit > AutoFit Window")],
    ["autoSave", "AutoSave", ribbon("Title bar > AutoSave toggle")],
    ["autoText", "AutoText", ribbon("Insert > Quick Parts > AutoText")],
    // --- B ---
    ["backgroundRemoval", "Background Removal", ribbon("Picture Format > Remove Background")],
    ["bibliography", "Bibliography", ribbon("References > Bibliography")],
    ["blankPage", "Blank Page", function () { insertBreak("Page"); }],
    ["blockAuthors", "Block Authors", ribbon("Review > Block Authors")],
    ["bold", "Bold", function () { toggleFont("bold"); }],
    ["bookmark", "Bookmark", shortcut("Ctrl+Shift+F5")],
    ["borderPainter", "Border Painter", ribbon("Table Design > Border Painter")],
    ["borders", "Borders", ribbon("Home > Borders")],
    ["bordersAndShading", "Borders and Shading", ribbon("Home > Borders > Borders and Shading")],
    ["breakLink", "Break Link", ribbon("Right-click link > Edit Link > Break Link")],
    ["breaks", "Breaks", ribbon("Layout > Breaks")],
    ["bringForward", "Bring Forward", ribbon("Shape Format > Bring Forward")],
    ["bringInFrontOfText", "Bring in Front of Text", ribbon("Shape Format > Wrap Text > In Front of Text")],
    ["bringToFront", "Bring to Front", ribbon("Shape Format > Bring to Front")],
    ["buildingBlocksOrganizer", "Building Blocks Organizer", ribbon("Insert > Quick Parts > Building Blocks Organizer")],
    ["bullets", "Bullets", ribbon("Home > Bullets")],

    // --- C ---
    ["calculate", "Calculate", ribbon("Status bar or Table Layout > Formula")],
    ["cancel", "Cancel", shortcut("Escape")],
    ["capitalizeEachWord", "Capitalize Each Word", ribbon("Home > Change Case > Capitalize Each Word")],
    ["caption", "Caption", ribbon("References > Insert Caption")],
    ["cellMargins", "Cell Margins", ribbon("Table Layout > Cell Margins")],
    ["changeCase", "Change Case", shortcut("Shift+F3")],
    ["changeChartType", "Change Chart Type", ribbon("Chart Design > Change Chart Type")],
    ["changePicture", "Change Picture", ribbon("Picture Format > Change Picture")],
    ["changeShape", "Change Shape", ribbon("Shape Format > Edit Shape > Change Shape")],
    ["characterBorder", "Character Border", ribbon("Home > Character Border")],
    ["characterShading", "Character Shading", ribbon("Home > Character Shading")],
    ["checkAccessibility", "Check Accessibility", ribbon("Review > Check Accessibility")],
    ["checkCompatibility", "Check Compatibility", ribbon("File > Info > Check for Issues > Check Compatibility")],
    ["citation", "Citation", ribbon("References > Insert Citation")],
    ["clearAllFormatting", "Clear All Formatting", function () { clearFormatting(); }],
    ["clipboard", "Clipboard", shortcut("Ctrl+C / Ctrl+V / Ctrl+X")],
    ["close", "Close", shortcut("Ctrl+W")],
    ["closeHeaderFooter", "Close Header and Footer", ribbon("Header & Footer > Close Header and Footer")],
    ["collapse", "Collapse", ribbon("View > Navigation Pane > Collapse")],
    ["collapseAll", "Collapse All", ribbon("View > Collapse All")],
    ["colorBlack", "Color: Black", function () { setFontColor("#000000"); }],
    ["colorBlue", "Color: Blue", function () { setFontColor("#0000FF"); }],
    ["colorBrown", "Color: Brown", function () { setFontColor("#993300"); }],
    ["colorCyan", "Color: Cyan", function () { setFontColor("#00FFFF"); }],
    ["colorDarkBlue", "Color: Dark Blue", function () { setFontColor("#000080"); }],
    ["colorDarkCyan", "Color: Dark Cyan", function () { setFontColor("#008080"); }],
    ["colorDarkGray", "Color: Dark Gray", function () { setFontColor("#808080"); }],
    ["colorDarkGreen", "Color: Dark Green", function () { setFontColor("#006400"); }],
    ["colorDarkMagenta", "Color: Dark Magenta", function () { setFontColor("#800080"); }],
    ["colorDarkRed", "Color: Dark Red", function () { setFontColor("#8B0000"); }],
    ["colorGold", "Color: Gold", function () { setFontColor("#FFD700"); }],
    ["colorGray", "Color: Gray", function () { setFontColor("#C0C0C0"); }],
    ["colorGreen", "Color: Green", function () { setFontColor("#008000"); }],
    ["colorLightBlue", "Color: Light Blue", function () { setFontColor("#ADD8E6"); }],
    ["colorLightGreen", "Color: Light Green", function () { setFontColor("#90EE90"); }],
    ["colorLime", "Color: Lime", function () { setFontColor("#00FF00"); }],
    ["colorMagenta", "Color: Magenta", function () { setFontColor("#FF00FF"); }],
    ["colorOlive", "Color: Olive", function () { setFontColor("#808000"); }],
    ["colorOrange", "Color: Orange", function () { setFontColor("#FFA500"); }],
    ["colorPink", "Color: Pink", function () { setFontColor("#FFC0CB"); }],
    ["colorPurple", "Color: Purple", function () { setFontColor("#800080"); }],
    ["colorRed", "Color: Red", function () { setFontColor("#FF0000"); }],
    ["colorTeal", "Color: Teal", function () { setFontColor("#008080"); }],
    ["colorViolet", "Color: Violet", function () { setFontColor("#EE82EE"); }],
    ["colorWhite", "Color: White", function () { setFontColor("#FFFFFF"); }],
    ["colorYellow", "Color: Yellow", function () { setFontColor("#FFFF00"); }],
    ["columnBreak", "Column Break", function () { insertBreak("Column"); }],
    ["columns", "Columns", ribbon("Layout > Columns")],
    ["combineDocuments", "Combine Documents", ribbon("Review > Compare > Combine")],
    ["compare", "Compare", ribbon("Review > Compare")],
    ["compareDocuments", "Compare Documents", ribbon("Review > Compare > Compare")],
    ["compressPictures", "Compress Pictures", ribbon("Picture Format > Compress Pictures")],
    ["contentControlProperties", "Content Control Properties", ribbon("Developer > Properties")],
    ["continuousBreak", "Continuous Section Break", function () { insertBreak("SectionContinuous"); }],
    ["continueNumbering", "Continue Numbering", ribbon("Right-click list > Continue Numbering")],
    ["convertTableToText", "Convert Table to Text", ribbon("Table Layout > Convert to Text")],
    ["convertTextToTable", "Convert Text to Table", ribbon("Insert > Table > Convert Text to Table")],
    ["copy", "Copy", shortcut("Ctrl+C")],
    ["coverPage", "Cover Page", ribbon("Insert > Cover Page")],
    ["createAutoText", "Create AutoText", ribbon("Insert > Quick Parts > AutoText > Save Selection")],
    ["crossReference", "Cross-reference", ribbon("References > Cross-reference")],
    ["customMargins", "Custom Margins", ribbon("Layout > Margins > Custom Margins")],
    ["customizeKeyboard", "Customize Keyboard", ribbon("File > Options > Customize Ribbon > Customize")],
    ["customizeQuickAccessToolbar", "Customize Quick Access Toolbar", ribbon("File > Options > Quick Access Toolbar")],
    ["customizeRibbon", "Customize Ribbon", ribbon("File > Options > Customize Ribbon")],
    ["cut", "Cut", shortcut("Ctrl+X")],
    // --- D ---
    ["dataLabels", "Data Labels", ribbon("Chart Design > Add Chart Element > Data Labels")],
    ["dataTable", "Data Table", ribbon("Chart Design > Add Chart Element > Data Table")],
    ["decimalTab", "Decimal Tab", ribbon("Paragraph dialog > Tabs > Decimal")],
    ["decreaseFontSize", "Decrease Font Size", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("size");
        return ctx.sync().then(function () { s.font.size = Math.max(1, s.font.size - 1); return ctx.sync(); });
      }, "Font size decreased.");
      else shortcut("Ctrl+Shift+<")();
    }],
    ["decreaseIndent", "Decrease Indent", function () { changeIndent(-36); }],
    ["defineNewBullet", "Define New Bullet", ribbon("Home > Bullets > Define New Bullet")],
    ["defineNewListStyle", "Define New List Style", ribbon("Home > Multilevel List > Define New List Style")],
    ["defineNewMultilevelList", "Define New Multilevel List", ribbon("Home > Multilevel List > Define New Multilevel List")],
    ["defineNewNumberFormat", "Define New Number Format", ribbon("Home > Numbering > Define New Number Format")],
    ["delete", "Delete", shortcut("Delete key")],
    ["deleteAllComments", "Delete All Comments", ribbon("Review > Delete > Delete All Comments")],
    ["deleteComment", "Delete Comment", ribbon("Review > Delete")],
    ["deleteCells", "Delete Cells", ribbon("Table Layout > Delete > Delete Cells")],
    ["deleteColumns", "Delete Columns", ribbon("Table Layout > Delete > Delete Columns")],
    ["deleteRows", "Delete Rows", ribbon("Table Layout > Delete > Delete Rows")],
    ["deleteTable", "Delete Table", ribbon("Table Layout > Delete > Delete Table")],
    ["demote", "Demote", ribbon("Outlining > Demote")],
    ["demoteToBodyText", "Demote to Body Text", ribbon("Outlining > Demote to Body Text")],
    ["designMode", "Design Mode", ribbon("Developer > Design Mode")],
    ["differentFirstPage", "Different First Page", ribbon("Header & Footer > Different First Page")],
    ["differentOddEven", "Different Odd & Even Pages", ribbon("Header & Footer > Different Odd & Even Pages")],
    ["distributeColumns", "Distribute Columns", ribbon("Table Layout > Distribute Columns")],
    ["distributeRows", "Distribute Rows", ribbon("Table Layout > Distribute Rows")],
    ["documentInspector", "Document Inspector", ribbon("File > Info > Check for Issues > Inspect Document")],
    ["documentMap", "Document Map", ribbon("View > Navigation Pane")],
    ["documentProtection", "Document Protection", ribbon("Review > Restrict Editing")],
    ["dontHyphenate", "Don't Hyphenate", ribbon("Paragraph dialog > Line and Page Breaks > Don't hyphenate")],
    ["doubleStrikethrough", "Double Strikethrough", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("doubleStrikethrough");
        return ctx.sync().then(function () { s.font.doubleStrikethrough = !s.font.doubleStrikethrough; return ctx.sync(); });
      }, "Double strikethrough toggled.");
      else showToast("Home > Font dialog > Double strikethrough");
    }],
    ["doubleUnderline", "Double Underline", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("underline");
        return ctx.sync().then(function () { s.font.underline = s.font.underline === "Double" ? "None" : "Double"; return ctx.sync(); });
      }, "Double underline toggled.");
      else showToast("Home > Font dialog > Double underline");
    }],
    ["draftView", "Draft View", ribbon("View > Draft")],
    ["drawTable", "Draw Table", ribbon("Insert > Table > Draw Table")],
    ["drawTextBox", "Draw Text Box", ribbon("Insert > Text Box > Draw Text Box")],
    ["drawingCanvas", "Drawing Canvas", ribbon("Insert > Shapes > New Drawing Canvas")],
    ["dropCap", "Drop Cap", ribbon("Insert > Drop Cap")],

    // --- E ---
    ["editingRestrictions", "Editing Restrictions", ribbon("Review > Restrict Editing")],
    ["editor", "Editor", ribbon("Home > Editor")],
    ["effects", "Effects (Theme)", ribbon("Design > Effects")],
    ["email", "Email", ribbon("File > Share > Email")],
    ["embedFonts", "Embed Fonts", ribbon("File > Options > Save > Embed fonts")],
    ["emphasisMark", "Emphasis Mark", ribbon("Home > Font dialog > Emphasis mark")],
    ["encloseCharacters", "Enclose Characters", ribbon("Home > Enclose Characters")],
    ["encryptWithPassword", "Encrypt with Password", ribbon("File > Info > Protect Document > Encrypt with Password")],
    ["endnote", "Endnote", ribbon("References > Insert Endnote")],
    ["envelopes", "Envelopes", ribbon("Mailings > Envelopes")],
    ["equation", "Equation", ribbon("Insert > Equation")],
    ["eraseTable", "Erase (Table)", ribbon("Table Layout > Eraser")],
    ["errorBars", "Error Bars", ribbon("Chart Design > Add Chart Element > Error Bars")],
    ["evenPageBreak", "Even Page Section Break", function () { insertBreak("SectionEven"); }],
    ["expand", "Expand", ribbon("View > Navigation Pane > Expand")],
    ["expandAll", "Expand All", ribbon("View > Expand All")],
    ["exportPdf", "Export to PDF/XPS", ribbon("File > Export > Create PDF/XPS")],
    // --- F ---
    ["field", "Field", ribbon("Insert > Quick Parts > Field")],
    ["fieldCodes", "Field Codes", shortcut("Alt+F9")],
    ["fieldShading", "Field Shading", ribbon("File > Options > Advanced > Show field shading")],
    ["fillColor", "Fill Color", ribbon("Home > Shading / Shape Format > Shape Fill")],
    ["find", "Find", shortcut("Ctrl+F")],
    ["findAndReplace", "Find and Replace", shortcut("Ctrl+H")],
    ["findNext", "Find Next", shortcut("Ctrl+G or F5")],
    ["firstLineIndent", "First Line Indent", function () { setIndent("firstLineIndent", 36); }],
    ["flipHorizontal", "Flip Horizontal", ribbon("Shape Format > Rotate > Flip Horizontal")],
    ["flipVertical", "Flip Vertical", ribbon("Shape Format > Rotate > Flip Vertical")],
    ["focusMode", "Focus Mode", ribbon("View > Focus")],
    ["fontArial", "Font: Arial", function () { setFontName("Arial"); }],
    ["fontCalibri", "Font: Calibri", function () { setFontName("Calibri"); }],
    ["fontCambria", "Font: Cambria", function () { setFontName("Cambria"); }],
    ["fontComicSans", "Font: Comic Sans MS", function () { setFontName("Comic Sans MS"); }],
    ["fontConsolas", "Font: Consolas", function () { setFontName("Consolas"); }],
    ["fontCourier", "Font: Courier New", function () { setFontName("Courier New"); }],
    ["fontGaramond", "Font: Garamond", function () { setFontName("Garamond"); }],
    ["fontGeorgia", "Font: Georgia", function () { setFontName("Georgia"); }],
    ["fontImpact", "Font: Impact", function () { setFontName("Impact"); }],
    ["fontLucidaConsole", "Font: Lucida Console", function () { setFontName("Lucida Console"); }],
    ["fontPalatino", "Font: Palatino Linotype", function () { setFontName("Palatino Linotype"); }],
    ["fontSegoeUI", "Font: Segoe UI", function () { setFontName("Segoe UI"); }],
    ["fontTahoma", "Font: Tahoma", function () { setFontName("Tahoma"); }],
    ["fontTimesNewRoman", "Font: Times New Roman", function () { setFontName("Times New Roman"); }],
    ["fontTrebuchet", "Font: Trebuchet MS", function () { setFontName("Trebuchet MS"); }],
    ["fontVerdana", "Font: Verdana", function () { setFontName("Verdana"); }],
    ["fontColor", "Font Color", ribbon("Home > Font Color dropdown")],
    ["fontDialog", "Font Dialog", shortcut("Ctrl+D")],
    ["fontSize8", "Font Size: 8", function () { setFontSize(8); }],
    ["fontSize9", "Font Size: 9", function () { setFontSize(9); }],
    ["fontSize10", "Font Size: 10", function () { setFontSize(10); }],
    ["fontSize10half", "Font Size: 10.5", function () { setFontSize(10.5); }],
    ["fontSize11", "Font Size: 11", function () { setFontSize(11); }],
    ["fontSize12", "Font Size: 12", function () { setFontSize(12); }],
    ["fontSize14", "Font Size: 14", function () { setFontSize(14); }],
    ["fontSize16", "Font Size: 16", function () { setFontSize(16); }],
    ["fontSize18", "Font Size: 18", function () { setFontSize(18); }],
    ["fontSize20", "Font Size: 20", function () { setFontSize(20); }],
    ["fontSize22", "Font Size: 22", function () { setFontSize(22); }],
    ["fontSize24", "Font Size: 24", function () { setFontSize(24); }],
    ["fontSize26", "Font Size: 26", function () { setFontSize(26); }],
    ["fontSize28", "Font Size: 28", function () { setFontSize(28); }],
    ["fontSize36", "Font Size: 36", function () { setFontSize(36); }],
    ["fontSize48", "Font Size: 48", function () { setFontSize(48); }],
    ["fontSize72", "Font Size: 72", function () { setFontSize(72); }],
    ["footer", "Footer", ribbon("Insert > Footer")],
    ["footerFromBottom", "Footer from Bottom", ribbon("Header & Footer > Footer from Bottom")],
    ["footnote", "Footnote", ribbon("References > Insert Footnote")],
    ["formatColumns", "Format Columns", ribbon("Layout > Columns > More Columns")],
    ["formatObject", "Format Object", ribbon("Right-click object > Format Object")],
    ["formatPageNumbers", "Format Page Numbers", ribbon("Insert > Page Number > Format Page Numbers")],
    ["formatPainter", "Format Painter", shortcut("Ctrl+Shift+C to copy, Ctrl+Shift+V to paste format")],
    ["formatTextEffects", "Format Text Effects", ribbon("Home > Font dialog > Text Effects")],
    ["formattingMarks", "Formatting Marks", shortcut("Ctrl+Shift+8")],
    ["formulas", "Formulas (Table)", ribbon("Table Layout > Formula")],
    ["fullScreenReading", "Full Screen Reading", ribbon("View > Read Mode")],

    // --- G ---
    ["goBack", "Go Back", shortcut("Shift+F5")],
    ["goTo", "Go To", shortcut("Ctrl+G or F5")],
    ["goToBookmark", "Go to Bookmark", shortcut("Ctrl+Shift+F5")],
    ["goToComment", "Go to Comment", ribbon("Review > Next / Previous comment")],
    ["goToEndnote", "Go to Endnote", ribbon("References > Next Endnote")],
    ["goToFooter", "Go to Footer", ribbon("Insert > Footer > Edit Footer")],
    ["goToFootnote", "Go to Footnote", ribbon("References > Next Footnote")],
    ["goToHeader", "Go to Header", ribbon("Insert > Header > Edit Header")],
    ["goToNextComment", "Go to Next Comment", ribbon("Review > Next")],
    ["goToNextSection", "Go to Next Section", ribbon("Header & Footer > Next Section")],
    ["goToPage", "Go to Page", shortcut("Ctrl+G")],
    ["goToPrevComment", "Go to Previous Comment", ribbon("Review > Previous")],
    ["goToPrevSection", "Go to Previous Section", ribbon("Header & Footer > Previous Section")],
    ["greetingLine", "Greeting Line", ribbon("Mailings > Greeting Line")],
    ["gridlines", "Gridlines", ribbon("View > Gridlines")],
    ["group", "Group", ribbon("Shape Format > Group > Group")],
    ["growFont", "Grow Font", shortcut("Ctrl+Shift+>")],
    // --- H ---
    ["hangingIndent", "Hanging Indent", function () { setIndent("firstLineIndent", -36); }],
    ["header", "Header", ribbon("Insert > Header")],
    ["headerFromTop", "Header from Top", ribbon("Header & Footer > Header from Top")],
    ["headerRow", "Header Row (Table)", ribbon("Table Design > Header Row")],
    ["heading1", "Heading 1", function () { setStyle("Heading 1"); }],
    ["heading2", "Heading 2", function () { setStyle("Heading 2"); }],
    ["heading3", "Heading 3", function () { setStyle("Heading 3"); }],
    ["heading4", "Heading 4", function () { setStyle("Heading 4"); }],
    ["heading5", "Heading 5", function () { setStyle("Heading 5"); }],
    ["heading6", "Heading 6", function () { setStyle("Heading 6"); }],
    ["headingRowsRepeat", "Heading Rows Repeat", ribbon("Table Layout > Repeat Header Rows")],
    ["help", "Help", shortcut("F1")],
    ["hidden", "Hidden Text", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("hidden");
        return ctx.sync().then(function () { s.font.hidden = !s.font.hidden; return ctx.sync(); });
      }, "Hidden text toggled.");
      else showToast("Home > Font dialog > Hidden");
    }],
    ["hideGrammaticalErrors", "Hide Grammatical Errors", ribbon("File > Options > Proofing")],
    ["hideSpellingErrors", "Hide Spelling Errors", ribbon("File > Options > Proofing")],
    ["highlightBlue", "Highlight: Blue", function () { setHighlight("#0000FF"); }],
    ["highlightBrightGreen", "Highlight: Bright Green", function () { setHighlight("#00FF00"); }],
    ["highlightCyan", "Highlight: Cyan", function () { setHighlight("#00FFFF"); }],
    ["highlightDarkBlue", "Highlight: Dark Blue", function () { setHighlight("#00008B"); }],
    ["highlightDarkRed", "Highlight: Dark Red", function () { setHighlight("#8B0000"); }],
    ["highlightDarkYellow", "Highlight: Dark Yellow", function () { setHighlight("#808000"); }],
    ["highlightGray25", "Highlight: Gray 25%", function () { setHighlight("#C0C0C0"); }],
    ["highlightGray50", "Highlight: Gray 50%", function () { setHighlight("#808080"); }],
    ["highlightGreen", "Highlight: Green", function () { setHighlight("#008000"); }],
    ["highlightMagenta", "Highlight: Magenta", function () { setHighlight("#FF00FF"); }],
    ["highlightNone", "Highlight: None (Remove)", function () { setHighlight(null); }],
    ["highlightPink", "Highlight: Pink", function () { setHighlight("#FF69B4"); }],
    ["highlightRed", "Highlight: Red", function () { setHighlight("#FF0000"); }],
    ["highlightTeal", "Highlight: Teal", function () { setHighlight("#008080"); }],
    ["highlightTurquoise", "Highlight: Turquoise", function () { setHighlight("#40E0D0"); }],
    ["highlightViolet", "Highlight: Violet", function () { setHighlight("#EE82EE"); }],
    ["highlightWhite", "Highlight: White", function () { setHighlight("#FFFFFF"); }],
    ["highlightYellow", "Highlight: Yellow", function () { setHighlight("#FFFF00"); }],
    ["hyphenation", "Hyphenation", ribbon("Layout > Hyphenation")],
    ["hyperlink", "Hyperlink", shortcut("Ctrl+K")],

    // --- I ---
    ["icons", "Icons", ribbon("Insert > Icons")],
    ["ignore", "Ignore (Spelling)", ribbon("Right-click > Ignore")],
    ["ignoreAll", "Ignore All (Spelling)", ribbon("Right-click > Ignore All")],
    ["immersiveReader", "Immersive Reader", ribbon("View > Immersive Reader")],
    ["importTheme", "Import Theme", ribbon("Design > Themes > Browse for Themes")],
    ["increaseFontSize", "Increase Font Size", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("size");
        return ctx.sync().then(function () { s.font.size = s.font.size + 1; return ctx.sync(); });
      }, "Font size increased.");
      else shortcut("Ctrl+Shift+>")();
    }],
    ["increaseIndent", "Increase Indent", function () { changeIndent(36); }],
    ["index", "Index", ribbon("References > Insert Index")],
    ["info", "Info", ribbon("File > Info")],
    ["insertAddressBlock", "Insert Address Block", ribbon("Mailings > Address Block")],
    ["insertCaption", "Insert Caption", ribbon("References > Insert Caption")],
    ["insertCells", "Insert Cells", ribbon("Table Layout > Insert > Insert Cells")],
    ["insertCitation", "Insert Citation", ribbon("References > Insert Citation")],
    ["insertColumnsLeft", "Insert Columns to the Left", ribbon("Table Layout > Insert Left")],
    ["insertColumnsRight", "Insert Columns to the Right", ribbon("Table Layout > Insert Right")],
    ["insertComment", "Insert Comment", shortcut("Ctrl+Alt+M")],
    ["insertDate", "Insert Date", ribbon("Insert > Date & Time")],
    ["insertDateTime", "Insert Date and Time", ribbon("Insert > Date & Time")],
    ["insertEndnote", "Insert Endnote", shortcut("Ctrl+Alt+D")],
    ["insertEquation", "Insert Equation", ribbon("Insert > Equation")],
    ["insertField", "Insert Field", ribbon("Insert > Quick Parts > Field")],
    ["insertFile", "Insert File", ribbon("Insert > Object > Text from File")],
    ["insertFootnote", "Insert Footnote", shortcut("Ctrl+Alt+F")],
    ["insertFrame", "Insert Frame", ribbon("Developer > Legacy Tools > Insert Frame")],
    ["insertHorizontalLine", "Insert Horizontal Line", function () { insertHtml('<hr style="border:1px solid #999;width:100%">', "Line inserted."); }],
    ["insertLeft", "Insert Left (Table)", ribbon("Table Layout > Insert Left")],
    ["insertMergeField", "Insert Merge Field", ribbon("Mailings > Insert Merge Field")],
    ["insertOnlinePictures", "Insert Online Pictures", ribbon("Insert > Online Pictures")],
    ["insertOnlineVideo", "Insert Online Video", ribbon("Insert > Online Video")],
    ["insertPageBreak", "Insert Page Break", function () { insertBreak("Page"); }],
    ["insertPageNumber", "Insert Page Number", ribbon("Insert > Page Number")],
    ["insertPicture", "Insert Picture", ribbon("Insert > Pictures")],
    ["insertRight", "Insert Right (Table)", ribbon("Table Layout > Insert Right")],
    ["insertRowsAbove", "Insert Rows Above", ribbon("Table Layout > Insert Above")],
    ["insertRowsBelow", "Insert Rows Below", ribbon("Table Layout > Insert Below")],
    ["insertSignatureLine", "Insert Signature Line", ribbon("Insert > Signature Line")],
    ["insertTable", "Insert Table", function () {
      if (isWord()) wordRun(function (ctx) {
        ctx.document.getSelection().insertTable(3, 3, "After", [["","",""],["","",""],["","",""]]);
        return ctx.sync();
      }, "Table inserted.");
      else if (isExcel()) excelRun(function (ctx) {
        var sh = ctx.workbook.worksheets.getActiveWorksheet();
        var r = ctx.workbook.getSelectedRange(); r.load("address");
        return ctx.sync().then(function () { sh.tables.add(r.address, true).name = "Table_" + Date.now(); return ctx.sync(); });
      }, "Table created.");
      else showToast("Not available.");
    }],
    ["insertTableOfAuthorities", "Insert Table of Authorities", ribbon("References > Insert Table of Authorities")],
    ["insertTableOfContents", "Insert Table of Contents", ribbon("References > Table of Contents")],
    ["insertTableOfFigures", "Insert Table of Figures", ribbon("References > Insert Table of Figures")],
    ["insertTextFromFile", "Insert Text from File", ribbon("Insert > Object > Text from File")],
    ["insertTime", "Insert Time", ribbon("Insert > Date & Time")],
    ["insertWordField", "Insert Word Field", ribbon("Mailings > Rules")],
    ["italic", "Italic", function () { toggleFont("italic"); }],
    // --- J-K ---
    ["justify", "Justify", function () { setAlignment("Justified"); }],
    ["keepLinesTogether", "Keep Lines Together", ribbon("Paragraph dialog > Line and Page Breaks > Keep lines together")],
    ["keepWithNext", "Keep with Next", ribbon("Paragraph dialog > Line and Page Breaks > Keep with next")],
    ["kerning", "Kerning", ribbon("Home > Font dialog > Advanced > Kerning")],

    // --- L ---
    ["labels", "Labels", ribbon("Mailings > Labels")],
    ["landscape", "Landscape", ribbon("Layout > Orientation > Landscape")],
    ["language", "Language", ribbon("Review > Language")],
    ["languagePreferences", "Language Preferences", ribbon("File > Options > Language")],
    ["lastColumn", "Last Column (Table)", ribbon("Table Design > Last Column")],
    ["launchDictation", "Launch Dictation", ribbon("Home > Dictate")],
    ["layoutOptions", "Layout Options", ribbon("Click object > Layout Options icon")],
    ["leftIndent", "Left Indent", function () { changeIndent(36); }],
    ["leftTab", "Left Tab", ribbon("Ruler > Click to set left tab")],
    ["legend", "Legend", ribbon("Chart Design > Add Chart Element > Legend")],
    ["ligatures", "Ligatures", ribbon("Home > Font dialog > Advanced > Ligatures")],
    ["limitFormatting", "Limit Formatting to a Selection of Styles", ribbon("Review > Restrict Editing > Formatting restrictions")],
    ["lineAndPageBreaks", "Line and Page Breaks", ribbon("Paragraph dialog > Line and Page Breaks")],
    ["lineBreak", "Line Break", shortcut("Shift+Enter")],
    ["lineNumbers", "Line Numbers", ribbon("Layout > Line Numbers")],
    ["lineSpacingSingle", "Line Spacing: Single (1.0)", function () { setLineSpacing(12); }],
    ["lineSpacing115", "Line Spacing: 1.15", function () { setLineSpacing(13.8); }],
    ["lineSpacing15", "Line Spacing: 1.5", function () { setLineSpacing(18); }],
    ["lineSpacingDouble", "Line Spacing: Double (2.0)", function () { setLineSpacing(24); }],
    ["lineSpacing25", "Line Spacing: 2.5", function () { setLineSpacing(30); }],
    ["lineSpacing3", "Line Spacing: Triple (3.0)", function () { setLineSpacing(36); }],
    ["linkToPrevious", "Link to Previous", ribbon("Header & Footer > Link to Previous")],
    ["lockDocument", "Lock Document", ribbon("Review > Restrict Editing")],
    ["lockTracking", "Lock Tracking", ribbon("Review > Track Changes > Lock Tracking")],
    ["lowercase", "Lowercase", ribbon("Home > Change Case > lowercase")],

    // --- M ---
    ["macros", "Macros", shortcut("Alt+F8")],
    ["mailMerge", "Mail Merge", ribbon("Mailings > Start Mail Merge")],
    ["manageAddIns", "Manage Add-ins", ribbon("File > Options > Add-ins")],
    ["manageSources", "Manage Sources", ribbon("References > Manage Sources")],
    ["manageStyles", "Manage Styles", ribbon("Home > Styles pane > Manage Styles")],
    ["manualHyphenation", "Manual Hyphenation", ribbon("Layout > Hyphenation > Manual")],
    ["margins", "Margins", ribbon("Layout > Margins")],
    ["markCitation", "Mark Citation", ribbon("References > Mark Citation")],
    ["markIndexEntry", "Mark Index Entry", shortcut("Alt+Shift+X")],
    ["markTocEntry", "Mark Table of Contents Entry", ribbon("References > Add Text")],
    ["matchCase", "Match Case", shortcut("Ctrl+H > Match case")],
    ["maximize", "Maximize Window", ribbon("Title bar > Maximize")],
    ["mergeCells", "Merge Cells", ribbon("Table Layout > Merge Cells")],
    ["mergeDocuments", "Merge Documents", ribbon("Review > Compare > Combine")],
    ["mergeToEmail", "Merge to E-mail", ribbon("Mailings > Finish & Merge > Send E-mail")],
    ["mergeToNewDoc", "Merge to New Document", ribbon("Mailings > Finish & Merge > Edit Individual Documents")],
    ["mergeToPrinter", "Merge to Printer", ribbon("Mailings > Finish & Merge > Print Documents")],
    ["minimize", "Minimize Window", ribbon("Title bar > Minimize")],
    ["mirrorIndents", "Mirror Indents", ribbon("Layout > Paragraph dialog > Mirror indents")],
    ["mirrorMargins", "Mirror Margins", ribbon("Layout > Margins > Mirrored")],
    ["modifyStyle", "Modify Style", ribbon("Home > Styles > Right-click > Modify")],
    ["moreColumns", "More Columns", ribbon("Layout > Columns > More Columns")],
    ["moreSymbols", "More Symbols", ribbon("Insert > Symbol > More Symbols")],
    ["moreUnderlines", "More Underlines", ribbon("Home > Font dialog > Underline style")],
    ["moveDown", "Move Down (Outline)", ribbon("Outlining > Move Down")],
    ["moveUp", "Move Up (Outline)", ribbon("Outlining > Move Up")],
    ["multilevelList", "Multilevel List", ribbon("Home > Multilevel List")],

    // --- N ---
    ["navigationPane", "Navigation Pane", shortcut("Ctrl+F or View > Navigation Pane")],
    ["new", "New", shortcut("Ctrl+N")],
    ["newComment", "New Comment", shortcut("Ctrl+Alt+M")],
    ["newDocument", "New Document", shortcut("Ctrl+N")],
    ["newFromTemplate", "New from Template", ribbon("File > New")],
    ["newMacro", "New Macro", ribbon("Developer > Macros > Create")],
    ["newStyle", "New Style", ribbon("Home > Styles pane > New Style")],
    ["newWindow", "New Window", ribbon("View > New Window")],
    ["nextChange", "Next Change", ribbon("Review > Next")],
    ["nextComment", "Next Comment", ribbon("Review > Next")],
    ["nextFootnote", "Next Footnote", ribbon("References > Next Footnote")],
    ["nextPage", "Next Page", shortcut("Ctrl+Page Down")],
    ["nextPageBreak", "Next Page Section Break", function () { insertBreak("SectionNext"); }],
    ["nextRecord", "Next Record (Mail Merge)", ribbon("Mailings > Next Record")],
    ["noBorder", "No Border", ribbon("Home > Borders > No Border")],
    ["noSpacing", "No Spacing (Style)", function () { setStyle("No Spacing"); }],
    ["normal", "Normal (Style)", function () { setStyle("Normal"); }],
    ["numberForm", "Number Form", ribbon("Home > Font dialog > Advanced > Number form")],
    ["numberSpacing", "Number Spacing", ribbon("Home > Font dialog > Advanced > Number spacing")],
    ["numbering", "Numbering", ribbon("Home > Numbering")],
    // --- O ---
    ["object", "Object", ribbon("Insert > Object")],
    ["oddPageBreak", "Odd Page Section Break", function () { insertBreak("SectionOdd"); }],
    ["officeClipboard", "Office Clipboard", shortcut("Ctrl+C twice or Home > Clipboard launcher")],
    ["onlinePictures", "Online Pictures", ribbon("Insert > Online Pictures")],
    ["open", "Open", shortcut("Ctrl+O")],
    ["openInBrowser", "Open in Browser", ribbon("File > Info > Edit in Browser")],
    ["openRecent", "Open Recent", ribbon("File > Open > Recent")],
    ["options", "Options", ribbon("File > Options")],
    ["orientation", "Orientation", ribbon("Layout > Orientation")],
    ["outlineView", "Outline View", ribbon("View > Outline")],

    // --- P ---
    ["pageBackground", "Page Background", ribbon("Design > Page Color")],
    ["pageBorder", "Page Border", ribbon("Design > Page Borders")],
    ["pageBreak", "Page Break", function () { insertBreak("Page"); }],
    ["pageBreakBefore", "Page Break Before", ribbon("Paragraph dialog > Line and Page Breaks > Page break before")],
    ["pageColor", "Page Color", ribbon("Design > Page Color")],
    ["pageLayoutView", "Page Layout View", ribbon("View > Print Layout")],
    ["pageNumber", "Page Number", ribbon("Insert > Page Number")],
    ["pageNumberFormat", "Page Number Format", ribbon("Insert > Page Number > Format Page Numbers")],
    ["pageSetup", "Page Setup", ribbon("Layout > Page Setup launcher")],
    ["paragraph", "Paragraph Dialog", ribbon("Home > Paragraph launcher")],
    ["paragraphMarks", "Paragraph Marks", shortcut("Ctrl+Shift+8")],
    ["paragraphShading", "Paragraph Shading", ribbon("Home > Shading")],
    ["paragraphSpacingAfter0", "Paragraph Spacing After: 0 pt", function () { setSpaceAfter(0); }],
    ["paragraphSpacingAfter6", "Paragraph Spacing After: 6 pt", function () { setSpaceAfter(6); }],
    ["paragraphSpacingAfter8", "Paragraph Spacing After: 8 pt", function () { setSpaceAfter(8); }],
    ["paragraphSpacingAfter10", "Paragraph Spacing After: 10 pt", function () { setSpaceAfter(10); }],
    ["paragraphSpacingAfter12", "Paragraph Spacing After: 12 pt", function () { setSpaceAfter(12); }],
    ["paragraphSpacingAfter24", "Paragraph Spacing After: 24 pt", function () { setSpaceAfter(24); }],
    ["paragraphSpacingBefore0", "Paragraph Spacing Before: 0 pt", function () { setSpaceBefore(0); }],
    ["paragraphSpacingBefore6", "Paragraph Spacing Before: 6 pt", function () { setSpaceBefore(6); }],
    ["paragraphSpacingBefore12", "Paragraph Spacing Before: 12 pt", function () { setSpaceBefore(12); }],
    ["paragraphSpacingBefore24", "Paragraph Spacing Before: 24 pt", function () { setSpaceBefore(24); }],
    ["paste", "Paste", shortcut("Ctrl+V")],
    ["pasteAll", "Paste All", shortcut("Ctrl+V (from Clipboard pane)")],
    ["pasteAsHyperlink", "Paste as Hyperlink", ribbon("Home > Paste > Paste Special > Paste as Hyperlink")],
    ["pasteAsPicture", "Paste as Picture", ribbon("Home > Paste > Paste Special > Picture")],
    ["pasteAsUnformatted", "Paste as Unformatted Text", shortcut("Ctrl+Shift+V")],
    ["pasteSpecial", "Paste Special", shortcut("Ctrl+Alt+V")],
    ["pasteTextOnly", "Paste Text Only", shortcut("Ctrl+Shift+V")],
    ["phoneticGuide", "Phonetic Guide", ribbon("Home > Phonetic Guide")],
    ["pictureBorder", "Picture Border", ribbon("Picture Format > Picture Border")],
    ["pictureEffects", "Picture Effects", ribbon("Picture Format > Picture Effects")],
    ["pictureLayout", "Picture Layout", ribbon("Picture Format > Picture Layout")],
    ["pictureStyles", "Picture Styles", ribbon("Picture Format > Picture Styles gallery")],
    ["portrait", "Portrait", ribbon("Layout > Orientation > Portrait")],
    ["position", "Position", ribbon("Shape Format > Position")],
    ["presentOnline", "Present Online", ribbon("File > Share > Present Online")],
    ["previewResults", "Preview Results (Mail Merge)", ribbon("Mailings > Preview Results")],
    ["previousChange", "Previous Change", ribbon("Review > Previous")],
    ["previousComment", "Previous Comment", ribbon("Review > Previous")],
    ["previousRecord", "Previous Record (Mail Merge)", ribbon("Mailings > Previous Record")],
    ["print", "Print", shortcut("Ctrl+P")],
    ["printLayout", "Print Layout", ribbon("View > Print Layout")],
    ["printPreview", "Print Preview", shortcut("Ctrl+F2")],
    ["promote", "Promote (Outline)", ribbon("Outlining > Promote")],
    ["proofingLanguage", "Proofing Language", ribbon("Review > Language > Set Proofing Language")],
    ["properties", "Properties", ribbon("File > Info > Properties")],
    ["protectDocument", "Protect Document", ribbon("Review > Restrict Editing")],
    ["publish", "Publish", ribbon("File > Export")],

    // --- Q ---
    ["quickParts", "Quick Parts", ribbon("Insert > Quick Parts")],
    ["quickPrint", "Quick Print", ribbon("File > Print > Quick Print")],
    ["quickStyles", "Quick Styles", ribbon("Home > Styles gallery")],
    ["quickTables", "Quick Tables", ribbon("Insert > Table > Quick Tables")],
    ["quote", "Quote (Style)", function () { setStyle("Quote"); }],
    // --- R ---
    ["readMode", "Read Mode", ribbon("View > Read Mode")],
    ["readingHighlight", "Reading Highlight", ribbon("Home > Find > Reading Highlight")],
    ["reapplyStyle", "Reapply Style", ribbon("Home > Styles > Right-click > Reapply")],
    ["recentDocuments", "Recent Documents", ribbon("File > Open > Recent")],
    ["recordMacro", "Record Macro", ribbon("Developer > Record Macro or View > Macros > Record Macro")],
    ["redo", "Redo", shortcut("Ctrl+Y")],
    ["reject", "Reject Change", ribbon("Review > Reject")],
    ["rejectAllChanges", "Reject All Changes", ribbon("Review > Reject > Reject All Changes")],
    ["removeAllFormatting", "Remove All Formatting", function () { clearFormatting(); }],
    ["removeBackground", "Remove Background", ribbon("Picture Format > Remove Background")],
    ["removeContentControl", "Remove Content Control", ribbon("Right-click control > Remove Content Control")],
    ["removeHyperlink", "Remove Hyperlink", ribbon("Right-click link > Remove Hyperlink")],
    ["removeSpaceAfter", "Remove Space After Paragraph", function () { setSpaceAfter(0); }],
    ["removeSpaceBefore", "Remove Space Before Paragraph", function () { setSpaceBefore(0); }],
    ["removeTableOfContents", "Remove Table of Contents", ribbon("References > Table of Contents > Remove Table of Contents")],
    ["removeWatermark", "Remove Watermark", ribbon("Design > Watermark > Remove Watermark")],
    ["repeat", "Repeat", shortcut("Ctrl+Y or F4")],
    ["repeatHeaderRows", "Repeat Header Rows", ribbon("Table Layout > Repeat Header Rows")],
    ["replace", "Replace", shortcut("Ctrl+H")],
    ["replaceAll", "Replace All", shortcut("Ctrl+H > Replace All")],
    ["research", "Research", ribbon("Review > Research")],
    ["resetCharFormatting", "Reset Character Formatting", shortcut("Ctrl+Space")],
    ["resetGraphic", "Reset Graphic", ribbon("Picture Format > Reset Picture")],
    ["resetParFormatting", "Reset Paragraph Formatting", shortcut("Ctrl+Q")],
    ["resetPicture", "Reset Picture", ribbon("Picture Format > Reset Picture")],
    ["resetPictureSize", "Reset Picture Size", ribbon("Picture Format > Reset Picture > Reset Picture & Size")],
    ["restartNumbering", "Restart Numbering", ribbon("Right-click list > Restart at 1")],
    ["restrictEditing", "Restrict Editing", ribbon("Review > Restrict Editing")],
    ["restrictFormatting", "Restrict Formatting", ribbon("Review > Restrict Editing > Formatting restrictions")],
    ["reviewingPane", "Reviewing Pane", ribbon("Review > Reviewing Pane")],
    ["reviewingPaneHorizontal", "Reviewing Pane Horizontal", ribbon("Review > Reviewing Pane > Horizontal")],
    ["reviewingPaneVertical", "Reviewing Pane Vertical", ribbon("Review > Reviewing Pane > Vertical")],
    ["rightIndent", "Right Indent", ribbon("Layout > Paragraph > Right indent")],
    ["rightTab", "Right Tab", ribbon("Ruler > Right tab stop")],
    ["rotateLeft90", "Rotate Left 90", ribbon("Shape Format > Rotate > Rotate Left 90")],
    ["rotateRight90", "Rotate Right 90", ribbon("Shape Format > Rotate > Rotate Right 90")],
    ["ruler", "Ruler", ribbon("View > Ruler")],
    ["rules", "Rules (Mail Merge)", ribbon("Mailings > Rules")],

    // --- S ---
    ["save", "Save", function () {
      Office.context.document.save(Office.AsyncResultStatus || {}, function (r) {
        if (r.status === Office.AsyncResultStatus.Failed) showToast("Save failed: " + r.error.message);
        else showToast("Saved!");
      });
    }],
    ["saveAs", "Save As", shortcut("F12")],
    ["saveAsPdf", "Save As PDF", ribbon("File > Export > Create PDF/XPS")],
    ["saveAsTemplate", "Save As Template", ribbon("File > Save As > Word Template")],
    ["saveCurrentTheme", "Save Current Theme", ribbon("Design > Themes > Save Current Theme")],
    ["screenClipping", "Screen Clipping", ribbon("Insert > Screenshot > Screen Clipping")],
    ["screenshot", "Screenshot", ribbon("Insert > Screenshot")],
    ["search", "Search", shortcut("Ctrl+F")],
    ["selectAll", "Select All", function () {
      if (isWord()) wordRun(function (ctx) { ctx.document.body.getRange().select(); return ctx.sync(); }, "All selected.");
      else shortcut("Ctrl+A")();
    }],
    ["selectCell", "Select Cell", ribbon("Table Layout > Select > Select Cell")],
    ["selectColumn", "Select Column", ribbon("Table Layout > Select > Select Column")],
    ["selectRow", "Select Row", ribbon("Table Layout > Select > Select Row")],
    ["selectTable", "Select Table", ribbon("Table Layout > Select > Select Table")],
    ["selectionPane", "Selection Pane", ribbon("Shape Format > Selection Pane")],
    ["sendAsAttachment", "Send as Attachment", ribbon("File > Share > Email > Send as Attachment")],
    ["sendAsPdf", "Send as PDF", ribbon("File > Share > Email > Send as PDF")],
    ["sendBackward", "Send Backward", ribbon("Shape Format > Send Backward")],
    ["sendBehindText", "Send Behind Text", ribbon("Shape Format > Wrap Text > Behind Text")],
    ["sendToBack", "Send to Back", ribbon("Shape Format > Send to Back")],
    ["sentenceCase", "Sentence Case", ribbon("Home > Change Case > Sentence case.")],
    ["setLanguage", "Set Language", ribbon("Review > Language > Set Proofing Language")],
    ["setNumberingValue", "Set Numbering Value", ribbon("Right-click list > Set Numbering Value")],
    ["shading", "Shading", ribbon("Home > Shading")],
    ["shapeFill", "Shape Fill", ribbon("Shape Format > Shape Fill")],
    ["shapeHeight", "Shape Height", ribbon("Shape Format > Height")],
    ["shapeOutline", "Shape Outline", ribbon("Shape Format > Shape Outline")],
    ["shapeStyles", "Shape Styles", ribbon("Shape Format > Shape Styles gallery")],
    ["shapeWidth", "Shape Width", ribbon("Shape Format > Width")],
    ["shapes", "Shapes", ribbon("Insert > Shapes")],
    ["share", "Share", ribbon("File > Share")],
    ["showAll", "Show All", shortcut("Ctrl+Shift+8")],
    ["showComments", "Show Comments", ribbon("Review > Show Comments")],
    ["showHide", "Show/Hide Paragraph Marks", shortcut("Ctrl+Shift+8")],
    ["showMarkup", "Show Markup", ribbon("Review > Show Markup")],
    ["showNotes", "Show Notes", ribbon("References > Show Notes")],
    ["showRevisionsBalloons", "Show Revisions in Balloons", ribbon("Review > Show Markup > Balloons")],
    ["shrinkFont", "Shrink Font", shortcut("Ctrl+Shift+<")],
    ["shrinkOnePage", "Shrink One Page", ribbon("Print Preview > Shrink One Page")],
    ["simpleMarkup", "Simple Markup", ribbon("Review > Display for Review > Simple Markup")],
    ["smallCaps", "Small Caps", shortcut("Ctrl+Shift+K")],
    ["smartArt", "SmartArt", ribbon("Insert > SmartArt")],
    ["snapToGrid", "Snap to Grid", ribbon("Shape Format > Align > Snap to Grid")],
    ["sort", "Sort", ribbon("Home > Sort")],
    ["sortAscending", "Sort Ascending", ribbon("Home > Sort (A-Z)")],
    ["sortDescending", "Sort Descending", ribbon("Home > Sort (Z-A)")],
    ["spaceAfter", "Space After", ribbon("Layout > Spacing After")],
    ["spaceBefore", "Space Before", ribbon("Layout > Spacing Before")],
    ["specialCharacters", "Special Characters", ribbon("Insert > Symbol > More Symbols > Special Characters")],
    ["spelling", "Spelling and Grammar", shortcut("F7")],
    ["splitCells", "Split Cells", ribbon("Table Layout > Split Cells")],
    ["splitTable", "Split Table", ribbon("Table Layout > Split Table")],
    ["splitWindow", "Split Window", ribbon("View > Split")],
    ["startEnforcingProtection", "Start Enforcing Protection", ribbon("Review > Restrict Editing > Yes, Start Enforcing Protection")],
    ["startMailMerge", "Start Mail Merge", ribbon("Mailings > Start Mail Merge")],
    ["startTracking", "Start Tracking", ribbon("Review > Track Changes")],
    ["strikethrough", "Strikethrough", function () { toggleFont("strikethrough"); }],
    ["styleHeading1", "Style: Heading 1", function () { setStyle("Heading 1"); }],
    ["styleHeading2", "Style: Heading 2", function () { setStyle("Heading 2"); }],
    ["styleHeading3", "Style: Heading 3", function () { setStyle("Heading 3"); }],
    ["styleIntenseEmphasis", "Style: Intense Emphasis", function () { setStyle("Intense Emphasis"); }],
    ["styleIntenseQuote", "Style: Intense Quote", function () { setStyle("Intense Quote"); }],
    ["styleIntenseReference", "Style: Intense Reference", function () { setStyle("Intense Reference"); }],
    ["styleListParagraph", "Style: List Paragraph", function () { setStyle("List Paragraph"); }],
    ["styleNormal", "Style: Normal", function () { setStyle("Normal"); }],
    ["styleSubtitle", "Style: Subtitle", function () { setStyle("Subtitle"); }],
    ["styleSubtleEmphasis", "Style: Subtle Emphasis", function () { setStyle("Subtle Emphasis"); }],
    ["styleSubtleReference", "Style: Subtle Reference", function () { setStyle("Subtle Reference"); }],
    ["styleTitle", "Style: Title", function () { setStyle("Title"); }],
    ["stylesDialog", "Styles Dialog", ribbon("Home > Styles launcher")],
    ["stylisticSets", "Stylistic Sets", ribbon("Home > Font dialog > Advanced > Stylistic sets")],
    ["subscript", "Subscript", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("subscript");
        return ctx.sync().then(function () { s.font.subscript = !s.font.subscript; return ctx.sync(); });
      }, "Subscript toggled.");
      else showToast("Not available.");
    }],
    ["sumFormula", "SUM Formula", ribbon("Table Layout > Formula > =SUM(ABOVE)")],
    ["superscript", "Superscript", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("superscript");
        return ctx.sync().then(function () { s.font.superscript = !s.font.superscript; return ctx.sync(); });
      }, "Superscript toggled.");
      else showToast("Not available.");
    }],
    ["switchWindows", "Switch Windows", ribbon("View > Switch Windows")],
    ["symbol", "Symbol", ribbon("Insert > Symbol")],
    ["symbolDialog", "Symbol Dialog", ribbon("Insert > Symbol > More Symbols")],
    ["synonyms", "Synonyms", ribbon("Right-click > Synonyms or Review > Thesaurus")],
    // --- T ---
    ["tableDesign", "Table Design", ribbon("Table Design tab")],
    ["tableGridlines", "Table Gridlines", ribbon("Table Layout > View Gridlines")],
    ["tableLayout", "Table Layout", ribbon("Table Layout tab")],
    ["tableOfAuthorities", "Table of Authorities", ribbon("References > Insert Table of Authorities")],
    ["tableOfContents", "Table of Contents", ribbon("References > Table of Contents")],
    ["tableOfFigures", "Table of Figures", ribbon("References > Insert Table of Figures")],
    ["tableProperties", "Table Properties", ribbon("Table Layout > Properties")],
    ["tableStyles", "Table Styles", ribbon("Table Design > Table Styles gallery")],
    ["tabs", "Tabs", ribbon("Paragraph dialog > Tabs")],
    ["templates", "Templates", ribbon("Developer > Document Template")],
    ["textBox", "Text Box", ribbon("Insert > Text Box")],
    ["textDirection", "Text Direction", ribbon("Table Layout > Text Direction")],
    ["textEffects", "Text Effects and Typography", ribbon("Home > Text Effects and Typography")],
    ["textFill", "Text Fill", ribbon("WordArt Format > Text Fill")],
    ["textHighlightColor", "Text Highlight Color", ribbon("Home > Text Highlight Color")],
    ["textOutline", "Text Outline", ribbon("WordArt Format > Text Outline")],
    ["textToSpeech", "Text to Speech", ribbon("Review > Read Aloud")],
    ["textWrapping", "Text Wrapping", ribbon("Shape Format > Wrap Text")],
    ["themeColors", "Theme Colors", ribbon("Design > Colors")],
    ["themeEffects", "Theme Effects", ribbon("Design > Effects")],
    ["themeFonts", "Theme Fonts", ribbon("Design > Fonts")],
    ["themes", "Themes", ribbon("Design > Themes")],
    ["thesaurus", "Thesaurus", shortcut("Shift+F7")],
    ["titleCase", "Title Case", ribbon("Home > Change Case > Capitalize Each Word")],
    ["toggleCase", "Toggle Case", ribbon("Home > Change Case > tOGGLE cASE")],
    ["toggleFieldCodes", "Toggle Field Codes", shortcut("Alt+F9")],
    ["toggleFullScreen", "Toggle Full Screen", shortcut("Alt+F11 or View > Read Mode")],
    ["totalRow", "Total Row (Table)", ribbon("Table Design > Total Row")],
    ["trackChanges", "Track Changes", shortcut("Ctrl+Shift+E")],
    ["trackChangesOptions", "Track Changes Options", ribbon("Review > Track Changes > Change Tracking Options")],
    ["translate", "Translate", ribbon("Review > Translate")],
    ["translateDocument", "Translate Document", ribbon("Review > Translate > Translate Document")],
    ["translateSelectedText", "Translate Selected Text", ribbon("Review > Translate > Translate Selection")],
    ["trendline", "Trendline", ribbon("Chart Design > Add Chart Element > Trendline")],
    ["trustCenterSettings", "Trust Center Settings", ribbon("File > Options > Trust Center > Trust Center Settings")],
    ["twoPages", "Two Pages", ribbon("View > Multiple Pages")],

    // --- U ---
    ["underline", "Underline", function () {
      if (isWord()) wordRun(function (ctx) {
        var s = ctx.document.getSelection(); s.font.load("underline");
        return ctx.sync().then(function () { s.font.underline = s.font.underline === "None" ? "Single" : "None"; return ctx.sync(); });
      }, "Underline toggled.");
      else toggleFont("underline");
    }],
    ["underlineColor", "Underline Color", ribbon("Home > Font dialog > Underline color")],
    ["underlineStyle", "Underline Style", ribbon("Home > Font dialog > Underline style")],
    ["undo", "Undo", shortcut("Ctrl+Z")],
    ["ungroup", "Ungroup", ribbon("Shape Format > Group > Ungroup")],
    ["updateAllFields", "Update All Fields", shortcut("Ctrl+A then F9")],
    ["updateBibliography", "Update Bibliography", ribbon("References > Update Citations and Bibliography")],
    ["updateField", "Update Field", shortcut("F9")],
    ["updateIndex", "Update Index", ribbon("References > Update Index")],
    ["updateLabels", "Update Labels", ribbon("Mailings > Update Labels")],
    ["updateStyleToMatch", "Update Style to Match Selection", ribbon("Home > Styles > Right-click > Update to Match Selection")],
    ["updateTableOfContents", "Update Table of Contents", ribbon("References > Update Table")],
    ["updateTableOfFigures", "Update Table of Figures", ribbon("References > Update Table")],
    ["uppercase", "UPPERCASE", ribbon("Home > Change Case > UPPERCASE")],

    // --- V ---
    ["verticalAlignment", "Vertical Alignment", ribbon("Layout > Page Setup > Layout tab > Vertical alignment")],
    ["viewCode", "View Code (VBA)", shortcut("Alt+F11")],
    ["viewFieldCodes", "View Field Codes", shortcut("Alt+F9")],
    ["viewFootnotes", "View Footnotes", ribbon("References > Show Notes")],
    ["viewGridlines", "View Gridlines", ribbon("View > Gridlines")],
    ["viewHeader", "View Header", ribbon("Insert > Header > Edit Header")],
    ["viewMacros", "View Macros", shortcut("Alt+F8")],
    ["viewMergedData", "View Merged Data", ribbon("Mailings > Preview Results")],
    ["viewRuler", "View Ruler", ribbon("View > Ruler")],
    ["viewSideBySide", "View Side by Side", ribbon("View > View Side by Side")],
    ["visualBasic", "Visual Basic", shortcut("Alt+F11")],

    // --- W ---
    ["watermark", "Watermark", ribbon("Design > Watermark")],
    ["webLayout", "Web Layout", ribbon("View > Web Layout")],
    ["wholePage", "Whole Page", ribbon("View > One Page")],
    ["widowOrphanControl", "Widow/Orphan Control", ribbon("Paragraph dialog > Line and Page Breaks > Widow/Orphan control")],
    ["wordArt", "WordArt", ribbon("Insert > WordArt")],
    ["wordArtStyles", "WordArt Styles", ribbon("WordArt Format > WordArt Styles gallery")],
    ["wordCount", "Word Count", shortcut("Ctrl+Shift+G or Review > Word Count")],
    ["wordOptions", "Word Options", ribbon("File > Options")],
    ["wordWrap", "Word Wrap", ribbon("Home > Paragraph > Allow line wrapping within cells")],
    ["wrapText", "Wrap Text", ribbon("Shape Format > Wrap Text")],

    // --- X ---
    ["xmlMappingPane", "XML Mapping Pane", ribbon("Developer > XML Mapping Pane")],
    ["xmlSchema", "XML Schema", ribbon("Developer > Schema")],

    // --- Z ---
    ["zoom100", "Zoom 100%", ribbon("View > Zoom > 100%")],
    ["zoomDialog", "Zoom Dialog", ribbon("View > Zoom")],
    ["zoomIn", "Zoom In", shortcut("Ctrl+Mouse wheel up or View > Zoom In")],
    ["zoomOnePage", "Zoom One Page", ribbon("View > One Page")],
    ["zoomOut", "Zoom Out", shortcut("Ctrl+Mouse wheel down or View > Zoom Out")],
    ["zoomPageWidth", "Zoom Page Width", ribbon("View > Page Width")],
    ["zoomTwoPages", "Zoom Two Pages", ribbon("View > Multiple Pages")],
  ];

  // ── Populate the dropdown ──────────────────────────────────────────
  function populateDropdown() {
    var sel = document.getElementById("command-list");
    sel.innerHTML = "";
    var placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.disabled = true;
    placeholder.selected = true;
    placeholder.textContent = "Choose a command (" + ALL_COMMANDS.length + " available)...";
    sel.appendChild(placeholder);

    ALL_COMMANDS.forEach(function (entry) {
      var opt = document.createElement("option");
      opt.value = entry[0];
      opt.textContent = entry[1];
      sel.appendChild(opt);
    });
  }

  // ── Build lookup map ───────────────────────────────────────────────
  var handlerMap = {};
  ALL_COMMANDS.forEach(function (entry) {
    handlerMap[entry[0]] = entry[2];
  });

  // ── Wire up Run button ─────────────────────────────────────────────
  function setup() {
    populateDropdown();
    var sel = document.getElementById("command-list");
    var btn = document.getElementById("run-btn");

    btn.addEventListener("click", function () {
      var cmd = sel.value;
      if (!cmd) { showToast("Pick a command first."); return; }
      var fn = handlerMap[cmd];
      if (fn) fn();
      else showToast("Unknown command: " + cmd);
    });

    // Also run on double-click of an option
    sel.addEventListener("dblclick", function () {
      var cmd = sel.value;
      if (cmd && handlerMap[cmd]) handlerMap[cmd]();
    });
  }

  // ── Initialize ─────────────────────────────────────────────────────
  Office.onReady(function () {
    setup();
  });
})();
