/*
 * SuperQAT – Task Pane (v3)
 * 2,199 Office 365 commands with search, category filter, and Office.js execution.
 */

/* global Office, Word, Excel, PowerPoint */

var CMD_DATA = require("../data/commands-slim.json");

(function () {
  "use strict";

  // ── Toast ────────────────────────────────────────────────────────────
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

  // ── Host detection ───────────────────────────────────────────────────
  function hostApp() {
    try { return Office.context.host; } catch (e) { return null; }
  }
  function isWord() { return hostApp() === Office.HostType.Word; }
  function isExcel() { return hostApp() === Office.HostType.Excel; }
  function isPowerPoint() { return hostApp() === Office.HostType.PowerPoint; }

  // ── Office.js helpers ────────────────────────────────────────────────
  function toggleFont(prop) {
    if (isWord()) {
      Word.run(function (c) {
        var s = c.document.getSelection(); s.font.load(prop);
        return c.sync().then(function () { s.font[prop] = !s.font[prop]; return c.sync(); });
      }).then(function () { showToast(prop.charAt(0).toUpperCase() + prop.slice(1) + " toggled."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) {
        var r = c.workbook.getSelectedRange(); r.format.font.load(prop);
        return c.sync().then(function () { r.format.font[prop] = !r.format.font[prop]; return c.sync(); });
      }).then(function () { showToast(prop.charAt(0).toUpperCase() + prop.slice(1) + " toggled."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontSize(sz) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.size = sz; return c.sync(); })
        .then(function () { showToast("Font size: " + sz); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.size = sz; return c.sync(); })
        .then(function () { showToast("Font size: " + sz); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontColor(color) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.color = color; return c.sync(); })
        .then(function () { showToast("Color applied."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.color = color; return c.sync(); })
        .then(function () { showToast("Color applied."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontName(name) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.name = name; return c.sync(); })
        .then(function () { showToast("Font: " + name); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.name = name; return c.sync(); })
        .then(function () { showToast("Font: " + name); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setAlignment(a) {
    if (isWord()) {
      Word.run(function (c) {
        var p = c.document.getSelection().paragraphs; p.load("items");
        return c.sync().then(function () { p.items.forEach(function (i) { i.alignment = a; }); return c.sync(); });
      }).then(function () { showToast("Alignment: " + a); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.horizontalAlignment = a; return c.sync(); })
        .then(function () { showToast("Alignment: " + a); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setLineSpacing(val) {
    if (isWord()) {
      Word.run(function (c) {
        var p = c.document.getSelection().paragraphs; p.load("items");
        return c.sync().then(function () { p.items.forEach(function (i) { i.lineSpacing = val; }); return c.sync(); });
      }).then(function () { showToast("Line spacing: " + val); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setHighlight(color) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.highlightColor = color; return c.sync(); })
        .then(function () { showToast(color ? "Highlighted." : "Highlight removed."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) {
        var r = c.workbook.getSelectedRange();
        if (color) { r.format.fill.color = color; } else { r.format.fill.clear(); }
        return c.sync();
      }).then(function () { showToast(color ? "Highlighted." : "Highlight removed."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Not available for this app."); }
  }

  function wordInsertBreak(type) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) { c.document.getSelection().insertBreak(type, "After"); return c.sync(); })
      .then(function () { showToast(type + " break inserted."); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }

  function setSpaceBefore(val) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) {
      var p = c.document.getSelection().paragraphs; p.load("items");
      return c.sync().then(function () { p.items.forEach(function (i) { i.spaceBefore = val; }); return c.sync(); });
    }).then(function () { showToast("Space before: " + val + "pt"); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }

  function setSpaceAfter(val) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) {
      var p = c.document.getSelection().paragraphs; p.load("items");
      return c.sync().then(function () { p.items.forEach(function (i) { i.spaceAfter = val; }); return c.sync(); });
    }).then(function () { showToast("Space after: " + val + "pt"); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }

  function setIndent(side, val) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) {
      var p = c.document.getSelection().paragraphs; p.load("items");
      return c.sync().then(function () {
        p.items.forEach(function (i) {
          if (side === "left") i.leftIndent = val;
          else if (side === "right") i.rightIndent = val;
          else if (side === "first") i.firstLineIndent = val;
        });
        return c.sync();
      });
    }).then(function () { showToast("Indent set."); })
      .catch(function (e) { showToast("Error: " + e.message); });
  }

  function applyStyle(name) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().style = name; return c.sync(); })
        .then(function () { showToast(name + " applied."); })
        .catch(function (e) { showToast("Error: " + e.message); });
    } else { showToast("Use the ribbon: Home \u2192 Styles \u2192 " + name); }
  }

  function ribbon(tab, tip) { showToast("Use the ribbon: " + tab + (tip ? " \u2192 " + tip : "")); }
  function shortcut(keys) { showToast("Keyboard shortcut: " + keys); }

  // ── Direct Office.js implementations (idMso -> handler) ──────────────
  var IMPL = {
    Bold: function () { toggleFont("bold"); },
    Italic: function () { toggleFont("italic"); },
    Underline: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("underline"); return c.sync().then(function () { s.font.underline = s.font.underline === "None" ? "Single" : "None"; return c.sync(); }); })
          .then(function () { showToast("Underline toggled."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { toggleFont("underline"); }
    },
    Strikethrough: function () { toggleFont("strikethrough"); },
    Subscript: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("subscript"); return c.sync().then(function () { s.font.subscript = !s.font.subscript; return c.sync(); }); })
          .then(function () { showToast("Subscript toggled."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+="); }
    },
    Superscript: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("superscript"); return c.sync().then(function () { s.font.superscript = !s.font.superscript; return c.sync(); }); })
          .then(function () { showToast("Superscript toggled."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+Shift++"); }
    },
    DoubleStrikethrough: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("doubleStrikethrough"); return c.sync().then(function () { s.font.doubleStrikethrough = !s.font.doubleStrikethrough; return c.sync(); }); })
          .then(function () { showToast("Double strikethrough toggled."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { ribbon("Home", "Font dialog \u2192 Double Strikethrough"); }
    },
    AllCaps: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("allCaps"); return c.sync().then(function () { s.font.allCaps = !s.font.allCaps; return c.sync(); }); })
          .then(function () { showToast("All Caps toggled."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+Shift+A"); }
    },

    AlignLeft: function () { setAlignment("Left"); },
    AlignCenter: function () { setAlignment("Center"); },
    AlignRight: function () { setAlignment("Right"); },
    AlignJustify: function () { setAlignment("Justified"); },
    AlignDistributed: function () { setAlignment("Distributed"); },

    FontSizeIncrease: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("size"); return c.sync().then(function () { s.font.size = s.font.size + 1; return c.sync(); }); })
          .then(function () { showToast("Font size increased."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+Shift+>"); }
    },
    FontSizeDecrease: function () {
      if (isWord()) {
        Word.run(function (c) { var s = c.document.getSelection(); s.font.load("size"); return c.sync().then(function () { s.font.size = Math.max(1, s.font.size - 1); return c.sync(); }); })
          .then(function () { showToast("Font size decreased."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+Shift+<"); }
    },

    FontColorPicker: function () { ribbon("Home", "Font Color picker"); },
    TextHighlightColorPicker: function () { ribbon("Home", "Text Highlight Color"); },

    IndentIncrease: function () { setIndent("left", 36); },
    IndentDecrease: function () { setIndent("left", 0); },
    IndentIncreaseWord: function () { setIndent("left", 36); },
    IndentDecreaseWord: function () { setIndent("left", 0); },

    SpacePara1: function () { setLineSpacing(12); },
    SpacePara15: function () { setLineSpacing(18); },
    SpacePara2: function () { setLineSpacing(24); },
    LineSpacing: function () { ribbon("Home", "Line Spacing"); },

    SpaceBeforeIncrease: function () { setSpaceBefore(6); },
    SpaceAfterIncrease: function () { setSpaceAfter(6); },
    SpaceBeforeDecrease: function () { setSpaceBefore(0); },
    SpaceAfterDecrease: function () { setSpaceAfter(0); },

    FileSave: function () {
      Office.context.document.save(Office.AsyncResultStatus || {}, function (r) {
        if (r.status === Office.AsyncResultStatus.Failed) { showToast("Save failed: " + r.error.message); }
        else { showToast("Saved!"); }
      });
    },
    FileOpen: function () { shortcut("Ctrl+O"); },
    FileClose: function () { shortcut("Ctrl+W"); },
    FilePrint: function () { shortcut("Ctrl+P"); },
    FilePrintPreview: function () { shortcut("Ctrl+F2"); },
    FileSaveAs: function () { shortcut("F12"); },
    FileNew: function () { shortcut("Ctrl+N"); },

    Copy: function () { shortcut("Ctrl+C"); },
    Cut: function () { shortcut("Ctrl+X"); },
    Paste: function () { shortcut("Ctrl+V"); },
    PasteSpecial: function () { shortcut("Ctrl+Alt+V"); },
    Undo: function () { shortcut("Ctrl+Z"); },
    Redo: function () { shortcut("Ctrl+Y"); },
    SelectAll: function () {
      if (isWord()) {
        Word.run(function (c) { c.document.body.getRange().select(); return c.sync(); })
          .then(function () { showToast("All selected."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { shortcut("Ctrl+A"); }
    },

    Find: function () { shortcut("Ctrl+F"); },
    Replace: function () { shortcut("Ctrl+H"); },
    GoTo: function () { shortcut("Ctrl+G"); },

    InsertTable: function () {
      if (isWord()) {
        Word.run(function (c) { c.document.getSelection().insertTable(3, 3, "After", [["", "", ""], ["", "", ""], ["", "", ""]]); return c.sync(); })
          .then(function () { showToast("Table inserted."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else if (isExcel()) {
        Excel.run(function (c) { var sh = c.workbook.worksheets.getActiveWorksheet(); var r = c.workbook.getSelectedRange(); r.load("address"); return c.sync().then(function () { sh.tables.add(r.address, true).name = "QuickTable_" + Date.now(); return c.sync(); }); })
          .then(function () { showToast("Table created."); }).catch(function (e) { showToast("Error: " + e.message); });
      } else { showToast("Not available for this app."); }
    },

    InsertPageBreak: function () { wordInsertBreak("Page"); },
    InsertSectionBreakNextPage: function () { wordInsertBreak("SectionNext"); },
    InsertSectionBreakContinuous: function () { wordInsertBreak("SectionContinuous"); },
    InsertSectionBreakEvenPage: function () { wordInsertBreak("EvenPage"); },
    InsertSectionBreakOddPage: function () { wordInsertBreak("OddPage"); },
    InsertColumnBreak: function () { wordInsertBreak("Column"); },

    InsertComment: function () { shortcut("Ctrl+Alt+M"); },
    InsertFootnote: function () { shortcut("Ctrl+Alt+F"); },
    InsertEndnote: function () { shortcut("Ctrl+Alt+D"); },
    InsertHyperlink: function () { shortcut("Ctrl+K"); },
    InsertBookmark: function () { shortcut("Ctrl+Shift+F5"); },
    InsertEquation: function () { shortcut("Alt+="); },
    InsertSymbol: function () { ribbon("Insert", "Symbol"); },
    InsertDateTime: function () { ribbon("Insert", "Date & Time"); },

    StyleNormal: function () { applyStyle("Normal"); },
    StyleHeading1: function () { applyStyle("Heading 1"); },
    StyleHeading2: function () { applyStyle("Heading 2"); },
    StyleHeading3: function () { applyStyle("Heading 3"); },
    StyleHeading4: function () { applyStyle("Heading 4"); },
    StyleHeading5: function () { applyStyle("Heading 5"); },
    StyleTitle: function () { applyStyle("Title"); },
    StyleSubtitle: function () { applyStyle("Subtitle"); },
    StyleEmphasis: function () { applyStyle("Emphasis"); },
    StyleStrong: function () { applyStyle("Strong"); },
    StyleQuote: function () { applyStyle("Quote"); },
    StyleListParagraph: function () { applyStyle("List Paragraph"); },

    NumberingBulletDefault: function () { ribbon("Home", "Bullets"); },
    NumberingDefault: function () { ribbon("Home", "Numbering"); },
    MultilevelListGallery: function () { ribbon("Home", "Multilevel List"); },

    TableMergeCells: function () { ribbon("Table Layout", "Merge Cells"); },
    TableSplitCells: function () { ribbon("Table Layout", "Split Cells"); },
    TableInsertRowsAbove: function () { ribbon("Table Layout", "Insert Above"); },
    TableInsertRowsBelow: function () { ribbon("Table Layout", "Insert Below"); },
    TableInsertColumnsLeft: function () { ribbon("Table Layout", "Insert Left"); },
    TableInsertColumnsRight: function () { ribbon("Table Layout", "Insert Right"); },
    TableDeleteRows: function () { ribbon("Table Layout", "Delete \u2192 Rows"); },
    TableDeleteColumns: function () { ribbon("Table Layout", "Delete \u2192 Columns"); },
    TableDeleteTable: function () { ribbon("Table Layout", "Delete \u2192 Table"); },
    SortAscendingExcel: function () { ribbon("Data", "Sort A to Z"); },
    SortDescendingExcel: function () { ribbon("Data", "Sort Z to A"); },
    SortDialog: function () { ribbon("Home", "Sort"); },

    SpellingAndGrammar: function () { shortcut("F7"); },
    Thesaurus: function () { shortcut("Shift+F7"); },
    WordCount: function () { ribbon("Review", "Word Count"); },
    TrackChanges: function () { shortcut("Ctrl+Shift+E"); },
    AcceptChange: function () { ribbon("Review", "Accept"); },
    RejectChange: function () { ribbon("Review", "Reject"); },
    AcceptAllChangesInDoc: function () { ribbon("Review", "Accept \u2192 Accept All Changes"); },
    RejectAllChangesInDoc: function () { ribbon("Review", "Reject \u2192 Reject All Changes"); },
    TranslateDocument: function () { ribbon("Review", "Translate \u2192 Translate Document"); },
    TranslateSelection: function () { ribbon("Review", "Translate \u2192 Translate Selection"); },
    AccessibilityChecker: function () { ribbon("Review", "Check Accessibility"); },

    ViewPrintLayoutWord: function () { ribbon("View", "Print Layout"); },
    ViewDraftView: function () { ribbon("View", "Draft"); },
    ViewOutlineMaster: function () { ribbon("View", "Outline"); },
    ViewWebLayoutWord: function () { ribbon("View", "Web Layout"); },
    ViewReadModeWord: function () { ribbon("View", "Read Mode"); },
    ViewRulerWord: function () { ribbon("View", "Ruler"); },
    ViewNavigationPane: function () { shortcut("Ctrl+F"); },
    ViewZoomDialog: function () { ribbon("View", "Zoom"); },

    PageOrientationPortrait: function () { ribbon("Layout", "Orientation \u2192 Portrait"); },
    PageOrientationLandscape: function () { ribbon("Layout", "Orientation \u2192 Landscape"); },
    PageSizeGallery: function () { ribbon("Layout", "Size"); },
    PageMarginsGallery: function () { ribbon("Layout", "Margins"); },
    ColumnsDialog: function () { ribbon("Layout", "Columns"); },
    PageSetupDialog: function () { ribbon("Layout", "Page Setup launcher"); },

    WatermarkGallery: function () { ribbon("Design", "Watermark"); },
    PageColorPicker: function () { ribbon("Design", "Page Color"); },
    PageBorderAndShadingDialog: function () { ribbon("Design", "Page Borders"); },
    ThemesGallery: function () { ribbon("Design", "Themes"); },

    HeaderInsert: function () { ribbon("Insert", "Header"); },
    FooterInsert: function () { ribbon("Insert", "Footer"); },
    InsertPageNumber: function () { ribbon("Insert", "Page Number"); },
    InsertPicture: function () { ribbon("Insert", "Pictures"); },
    InsertOnlinePicture: function () { ribbon("Insert", "Online Pictures"); },
    SmartArtInsert: function () { ribbon("Insert", "SmartArt"); },
    ChartInsert: function () { ribbon("Insert", "Chart"); },
    InsertTextBox: function () { ribbon("Insert", "Text Box"); },
    InsertWordArt: function () { ribbon("Insert", "WordArt"); },
    InsertDropCap: function () { ribbon("Insert", "Drop Cap"); },
    InsertSignatureLine: function () { ribbon("Insert", "Signature Line"); },
    ObjectInsert: function () { ribbon("Insert", "Object"); },

    TableOfContentsInsert: function () { ribbon("References", "Table of Contents"); },
    TableOfContentsUpdate: function () { ribbon("References", "Update Table"); },
    InsertCitation: function () { ribbon("References", "Insert Citation"); },
    ManageSources: function () { ribbon("References", "Manage Sources"); },
    InsertCaption: function () { ribbon("References", "Insert Caption"); },
    InsertIndexDialog: function () { ribbon("References", "Insert Index"); },
    MarkIndexEntry: function () { shortcut("Alt+Shift+X"); },

    MailMergeStartMailMerge: function () { ribbon("Mailings", "Start Mail Merge"); },
    MailMergeSelectRecipients: function () { ribbon("Mailings", "Select Recipients"); },
    MailMergeInsertMergeField: function () { ribbon("Mailings", "Insert Merge Field"); },
    MailMergePreviewResults: function () { ribbon("Mailings", "Preview Results"); },
    MailMergeFinishAndMerge: function () { ribbon("Mailings", "Finish & Merge"); },

    MacroRecord: function () { ribbon("View", "Macros \u2192 Record Macro"); },
    MacroPlay: function () { shortcut("Alt+F8"); },
    VisualBasicEditor: function () { shortcut("Alt+F11"); },

    FormatPainter: function () { shortcut("Ctrl+Shift+C to copy, Ctrl+Shift+V to paste format"); },
    ClearAllFormatting: function () { shortcut("Ctrl+Space"); },

    ObjectBringToFront: function () { ribbon("Shape Format", "Bring to Front"); },
    ObjectSendToBack: function () { ribbon("Shape Format", "Send to Back"); },
    ObjectBringForward: function () { ribbon("Shape Format", "Bring Forward"); },
    ObjectSendBackward: function () { ribbon("Shape Format", "Send Backward"); },
    ObjectsGroup: function () { ribbon("Shape Format", "Group"); },
    ObjectsUngroup: function () { ribbon("Shape Format", "Ungroup"); },

    ShapeFillColorPicker: function () { ribbon("Shape Format", "Shape Fill"); },
    ShapeOutlineColorPicker: function () { ribbon("Shape Format", "Shape Outline"); },
    ShapeEffectsMenu: function () { ribbon("Shape Format", "Shape Effects"); },

    PictureCompress: function () { ribbon("Picture Format", "Compress Pictures"); },
    PictureCrop: function () { ribbon("Picture Format", "Crop"); },
    PictureResetAndSize: function () { ribbon("Picture Format", "Reset Picture"); },
    PictureChangeMenu: function () { ribbon("Picture Format", "Change Picture"); },

    FieldToggle: function () { shortcut("Alt+F9"); },
    UpdateFields: function () { shortcut("Ctrl+A then F9"); },

    Help: function () { shortcut("F1"); },
    About: function () { ribbon("File", "Account"); }
  };

  // ── Convert PascalCase idMso to readable label ───────────────────────
  function humanize(name) {
    if (name.charAt(0) === "_") name = name.substring(1);
    return name
      .replace(/([a-z])([A-Z])/g, "$1 $2")
      .replace(/([A-Z]+)([A-Z][a-z])/g, "$1 $2")
      .replace(/(\d+)/g, " $1 ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // ── Friendly tab names ───────────────────────────────────────────────
  var TAB_LABELS = {
    "TabHome": "Home",
    "TabInsert": "Insert",
    "TabPageLayoutWord": "Layout (Word)",
    "TabPageLayoutExcel": "Layout (Excel)",
    "TabReferences": "References",
    "TabMailings": "Mailings",
    "TabReview": "Review",
    "TabReviewWord": "Review (Word)",
    "TabView": "View",
    "TabData": "Data",
    "TabFormulas": "Formulas",
    "TabDeveloper": "Developer",
    "TabDesign": "Design",
    "TabWordDesign": "Design (Word)",
    "TabDrawInk": "Draw",
    "TabRecording": "Recording",
    "TabAnimations": "Animations",
    "TabTransitions": "Transitions",
    "TabSlideShow": "Slide Show",
    "TabSlideMaster": "Slide Master",
    "TabOutlining": "Outlining",
    "TabPrintPreview": "Print Preview",
    "TabAutomate": "Automate",
    "TabHelp": "Help",
    "HelpTab": "Help",
    "TabTableToolsLayout": "Table Layout",
    "TabTableToolsDesign": "Table Design",
    "TabTableToolsDesignExcel": "Table Design (Excel)",
    "TabPictureToolsFormat": "Picture Format",
    "TabTextBoxToolsFormat": "Text Box Format",
    "TabChartToolsDesignNew": "Chart Design",
    "TabChartToolsLayout": "Chart Layout",
    "TabSmartArtToolsDesign": "SmartArt Design",
    "TabSmartArtToolsFormat": "SmartArt Format",
    "TabHeaderAndFooterToolsDesign": "Header & Footer",
    "TabPivotTableToolsOptions": "PivotTable Analyze",
    "TabPivotTableToolsDesign": "PivotTable Design",
    "TabSparklineDesign": "Sparkline Design",
    "TabEquationToolsDesign": "Equation",
    "TabBackgroundRemoval": "Background Removal",
    "None (Not in the Ribbon)": "Not in Ribbon",
    "None (Context Menu)": "Context Menu",
    "Quick Access Toolbar": "Quick Access Toolbar",
    "Other": "Other"
  };

  function friendlyTab(raw) {
    return TAB_LABELS[raw] || raw.replace(/^Tab/, "").replace(/([a-z])([A-Z])/g, "$1 $2");
  }

  // ── Build the command list from slim data ────────────────────────────
  var TABS = CMD_DATA.tabs;
  var ALL_COMMANDS = CMD_DATA.cmds.map(function (c) {
    return {
      name: c[0],
      label: humanize(c[0]),
      tabIdx: c[1],
      tab: TABS[c[1]],
      apps: c[2]
    };
  });

  // ── Filter commands for current host app ─────────────────────────────
  function appBit() {
    if (isWord()) return 1;
    if (isExcel()) return 2;
    if (isPowerPoint()) return 4;
    return 7;
  }

  // ── Run a command ────────────────────────────────────────────────────
  function runCommand(cmd) {
    if (IMPL[cmd.name]) {
      IMPL[cmd.name]();
    } else {
      var tabLabel = friendlyTab(cmd.tab);
      showToast("Use the ribbon: " + tabLabel + " \u2192 " + cmd.label);
    }
  }

  // ── UI setup ─────────────────────────────────────────────────────────
  function setup() {
    var searchEl = document.getElementById("search");
    var tabFilterEl = document.getElementById("tab-filter");
    var selectEl = document.getElementById("command-list");
    var btnEl = document.getElementById("run-btn");
    var infoEl = document.getElementById("cmd-info");
    var countEl = document.getElementById("cmd-count");

    var bit = appBit();
    var appCommands = ALL_COMMANDS.filter(function (c) { return c.apps & bit; });

    var tabSet = {};
    appCommands.forEach(function (c) { tabSet[c.tab] = (tabSet[c.tab] || 0) + 1; });
    var tabEntries = Object.keys(tabSet).sort(function (a, b) {
      return tabSet[b] - tabSet[a];
    });
    tabFilterEl.innerHTML = '<option value="">All Categories (' + appCommands.length + ')</option>';
    tabEntries.forEach(function (t) {
      var opt = document.createElement("option");
      opt.value = t;
      opt.textContent = friendlyTab(t) + " (" + tabSet[t] + ")";
      tabFilterEl.appendChild(opt);
    });

    var visibleCommands = appCommands;

    function render() {
      var query = (searchEl.value || "").toLowerCase();
      var tabVal = tabFilterEl.value;

      visibleCommands = appCommands.filter(function (c) {
        if (tabVal && c.tab !== tabVal) return false;
        if (query && c.label.toLowerCase().indexOf(query) === -1 && c.name.toLowerCase().indexOf(query) === -1) return false;
        return true;
      });

      selectEl.innerHTML = "";
      visibleCommands.forEach(function (c, idx) {
        var opt = document.createElement("option");
        opt.value = idx;
        opt.textContent = (IMPL[c.name] ? "\u2713 " : "") + c.label;
        selectEl.appendChild(opt);
      });

      countEl.textContent = visibleCommands.length + " commands";
      infoEl.textContent = "";
    }

    render();

    searchEl.addEventListener("input", render);
    tabFilterEl.addEventListener("change", render);

    selectEl.addEventListener("change", function () {
      var idx = parseInt(selectEl.value, 10);
      if (isNaN(idx) || !visibleCommands[idx]) return;
      var cmd = visibleCommands[idx];
      var parts = [];
      parts.push(friendlyTab(cmd.tab));
      if (IMPL[cmd.name]) parts.push("\u2713 Direct execution");
      else parts.push("Ribbon guidance");
      var appNames = [];
      if (cmd.apps & 1) appNames.push("Word");
      if (cmd.apps & 2) appNames.push("Excel");
      if (cmd.apps & 4) appNames.push("PPT");
      parts.push(appNames.join("/"));
      infoEl.textContent = parts.join(" \u2022 ");
    });

    function executeSelected() {
      var idx = parseInt(selectEl.value, 10);
      if (isNaN(idx) || !visibleCommands[idx]) { showToast("Pick a command first."); return; }
      runCommand(visibleCommands[idx]);
    }

    btnEl.addEventListener("click", executeSelected);
    selectEl.addEventListener("dblclick", executeSelected);
  }

  Office.onReady(function () { setup(); });
})();
