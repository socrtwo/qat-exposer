/*
 * Office Quick Access Add-in - Task Pane
 * Single dropdown exposing every QAT-customizable command.
 */

/* global Office, Word, Excel, PowerPoint */

(function () {
  "use strict";

  // ── Toast notification ────────────────────────────────────────────
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

  // ── Office.js helpers ─────────────────────────────────────────────
  function isWord()  { return Office.context.host === Office.HostType.Word; }
  function isExcel() { return Office.context.host === Office.HostType.Excel; }

  function toggleFont(prop) {
    if (isWord()) {
      Word.run(function (c) {
        var s = c.document.getSelection(); s.font.load(prop);
        return c.sync().then(function () { s.font[prop] = !s.font[prop]; return c.sync(); });
      }).then(function () { showToast(prop.charAt(0).toUpperCase()+prop.slice(1)+" toggled."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) {
        var r = c.workbook.getSelectedRange(); r.format.font.load(prop);
        return c.sync().then(function () { r.format.font[prop] = !r.format.font[prop]; return c.sync(); });
      }).then(function () { showToast(prop.charAt(0).toUpperCase()+prop.slice(1)+" toggled."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontSize(sz) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.size = sz; return c.sync(); })
        .then(function () { showToast("Font size: "+sz); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.size = sz; return c.sync(); })
        .then(function () { showToast("Font size: "+sz); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontColor(color) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.color = color; return c.sync(); })
        .then(function () { showToast("Color applied."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.color = color; return c.sync(); })
        .then(function () { showToast("Color applied."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setFontName(name) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.name = name; return c.sync(); })
        .then(function () { showToast("Font: "+name); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.font.name = name; return c.sync(); })
        .then(function () { showToast("Font: "+name); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setAlignment(a) {
    if (isWord()) {
      Word.run(function (c) {
        var p = c.document.getSelection().paragraphs; p.load("items");
        return c.sync().then(function () { p.items.forEach(function (i) { i.alignment = a; }); return c.sync(); });
      }).then(function () { showToast("Alignment: "+a); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) { c.workbook.getSelectedRange().format.horizontalAlignment = a; return c.sync(); })
        .then(function () { showToast("Alignment: "+a); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setLineSpacing(val) {
    if (isWord()) {
      Word.run(function (c) {
        var p = c.document.getSelection().paragraphs; p.load("items");
        return c.sync().then(function () { p.items.forEach(function (i) { i.lineSpacing = val; }); return c.sync(); });
      }).then(function () { showToast("Line spacing: "+val); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function setHighlight(color) {
    if (isWord()) {
      Word.run(function (c) { c.document.getSelection().font.highlightColor = color; return c.sync(); })
        .then(function () { showToast(color ? "Highlighted." : "Highlight removed."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else if (isExcel()) {
      Excel.run(function (c) {
        var r = c.workbook.getSelectedRange();
        if (color) { r.format.fill.color = color; } else { r.format.fill.clear(); }
        return c.sync();
      }).then(function () { showToast(color ? "Highlighted." : "Highlight removed."); })
        .catch(function (e) { showToast("Error: "+e.message); });
    } else { showToast("Not available for this app."); }
  }

  function wordInsertBreak(type) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) { c.document.getSelection().insertBreak(type, "After"); return c.sync(); })
      .then(function () { showToast(type+" break inserted."); })
      .catch(function (e) { showToast("Error: "+e.message); });
  }

  function wordInsertHtml(html, msg) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) { c.document.getSelection().insertHtml(html, "After"); return c.sync(); })
      .then(function () { showToast(msg); })
      .catch(function (e) { showToast("Error: "+e.message); });
  }

  function setSpaceBefore(val) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) {
      var p = c.document.getSelection().paragraphs; p.load("items");
      return c.sync().then(function () { p.items.forEach(function (i) { i.spaceBefore = val; }); return c.sync(); });
    }).then(function () { showToast("Space before: "+val+"pt"); })
      .catch(function (e) { showToast("Error: "+e.message); });
  }

  function setSpaceAfter(val) {
    if (!isWord()) { showToast("Only available in Word."); return; }
    Word.run(function (c) {
      var p = c.document.getSelection().paragraphs; p.load("items");
      return c.sync().then(function () { p.items.forEach(function (i) { i.spaceAfter = val; }); return c.sync(); });
    }).then(function () { showToast("Space after: "+val+"pt"); })
      .catch(function (e) { showToast("Error: "+e.message); });
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
      .catch(function (e) { showToast("Error: "+e.message); });
  }

  function ribbon(tab, tip) { showToast("Use the ribbon: " + tab + (tip ? " \u2192 " + tip : "")); }
  function shortcut(keys) { showToast("Keyboard shortcut: " + keys); }

  // ── Master command list (alphabetical) ────────────────────────────
  // Each entry: [label, handler]
  // handler is a function that either executes via Office.js or shows guidance.

  var ALL_COMMANDS = [
    ["About", function(){ribbon("File","Account");}],
    ["Accept All Changes in Document", function(){ribbon("Review","Accept > Accept All Changes");}],
    ["Accept and Move to Next", function(){ribbon("Review","Accept");}],
    ["Accessibility Checker", function(){ribbon("Review","Check Accessibility");}],
    ["Add Chart Element", function(){ribbon("Chart Design","Add Chart Element");}],
    ["Add Horizontal Line", function(){wordInsertHtml('<hr style="border:1px solid #999;width:100%">',"Line inserted.");}],
    ["Add Shape", function(){ribbon("Insert","Shapes");}],
    ["Add Text (Table of Contents)", function(){ribbon("References","Add Text");}],
    ["Adjust List Indents", function(){ribbon("Home","Multilevel List > Adjust");}],
    ["Advanced Find", function(){shortcut("Ctrl+Shift+F");}],
    ["Align Bottom", function(){ribbon("Table Layout","Align Bottom");}],
    ["Align Bottom Center", function(){ribbon("Table Layout","Align Bottom Center");}],
    ["Align Bottom Left", function(){ribbon("Table Layout","Align Bottom Left");}],
    ["Align Bottom Right", function(){ribbon("Table Layout","Align Bottom Right");}],
    ["Align Center", function(){setAlignment("Center");}],
    ["Align Left", function(){setAlignment("Left");}],
    ["Align Middle Center", function(){ribbon("Table Layout","Align Center");}],
    ["Align Middle Left", function(){ribbon("Table Layout","Align Center Left");}],
    ["Align Middle Right", function(){ribbon("Table Layout","Align Center Right");}],
    ["Align Right", function(){setAlignment("Right");}],
    ["Align Top", function(){ribbon("Table Layout","Align Top");}],
    ["Align Top Center", function(){ribbon("Table Layout","Align Top Center");}],
    ["Align Top Left", function(){ribbon("Table Layout","Align Top Left");}],
    ["Align Top Right", function(){ribbon("Table Layout","Align Top Right");}],
    ["Arrange All Windows", function(){ribbon("View","Arrange All");}],
    ["Attach Template", function(){ribbon("Developer","Document Template");}],
    ["AutoCorrect Options", function(){ribbon("File","Options > Proofing > AutoCorrect Options");}],
    ["AutoFit Contents", function(){ribbon("Table Layout","AutoFit > AutoFit Contents");}],
    ["AutoFit Window", function(){ribbon("Table Layout","AutoFit > AutoFit Window");}],
    ["AutoFormat", function(){ribbon("File","Options > Proofing > AutoFormat");}],
    ["AutoSave", function(){ribbon("File","AutoSave toggle (top-left)");}],
    ["AutoShapes: Action Buttons", function(){ribbon("Insert","Shapes > Action Buttons");}],
    ["AutoShapes: Basic Shapes", function(){ribbon("Insert","Shapes > Basic Shapes");}],
    ["AutoShapes: Block Arrows", function(){ribbon("Insert","Shapes > Block Arrows");}],
    ["AutoShapes: Callouts", function(){ribbon("Insert","Shapes > Callouts");}],
    ["AutoShapes: Equation Shapes", function(){ribbon("Insert","Shapes > Equation Shapes");}],
    ["AutoShapes: Flowchart", function(){ribbon("Insert","Shapes > Flowchart");}],
    ["AutoShapes: Lines", function(){ribbon("Insert","Shapes > Lines");}],
    ["AutoShapes: Rectangles", function(){ribbon("Insert","Shapes > Rectangles");}],
    ["AutoShapes: Stars and Banners", function(){ribbon("Insert","Shapes > Stars and Banners");}],
    ["AutoText", function(){ribbon("Insert","Quick Parts > AutoText");}],
    ["Block Authors", function(){ribbon("Review","Block Authors");}],
    ["Block Authors: All My Edited Sections", function(){ribbon("Review","Block Authors > Block All My Edited Sections");}],
    ["Bold", function(){toggleFont("bold");}],
    ["Bookmark", function(){shortcut("Ctrl+Shift+F5");}],
    ["Borders and Shading", function(){ribbon("Home","Borders dropdown > Borders and Shading");}],
    ["Bring Forward", function(){ribbon("Shape Format","Bring Forward");}],
    ["Bring in Front of Text", function(){ribbon("Shape Format","Bring Forward > In Front of Text");}],
    ["Bring to Front", function(){ribbon("Shape Format","Bring to Front");}],
    ["Building Blocks Organizer", function(){ribbon("Insert","Quick Parts > Building Blocks Organizer");}],
    ["Bullets", function(){ribbon("Home","Bullets");}],
    ["Calculate", function(){ribbon("Review","Word Count / Formula");}],
    ["Cell Margins", function(){ribbon("Table Layout","Cell Margins");}],
    ["Center", function(){setAlignment("Center");}],
    ["Change Chart Type", function(){ribbon("Chart Design","Change Chart Type");}],
    ["Change Colors (Chart)", function(){ribbon("Chart Design","Change Colors");}],
    ["Change Picture", function(){ribbon("Picture Format","Change Picture");}],
    ["Change Shape", function(){ribbon("Shape Format","Edit Shape > Change Shape");}],
    ["Character Spacing: Condensed", function(){ribbon("Home","Font dialog > Advanced > Spacing > Condensed");}],
    ["Character Spacing: Expanded", function(){ribbon("Home","Font dialog > Advanced > Spacing > Expanded");}],
    ["Check Compatibility", function(){ribbon("File","Info > Check for Issues > Compatibility");}],
    ["Clipboard Pane", function(){shortcut("Ctrl+C twice (or Home > Clipboard launcher)");}],
    ["Close", function(){shortcut("Ctrl+W");}],
    ["Close All", function(){ribbon("File","Close All");}],
    ["Close Header and Footer", function(){ribbon("Header & Footer","Close Header and Footer");}],
    ["Close Outline View", function(){ribbon("Outlining","Close Outline View");}],
    ["Collapse All Headings", function(){ribbon("View","Outline > Collapse All");}],
    ["Collapse Heading", function(){ribbon("Home","Right-click heading > Expand/Collapse > Collapse Heading");}],
    ["Collapse Subdocuments", function(){ribbon("Outlining","Master Document > Collapse Subdocuments");}],
    ["Color: Aqua", function(){setFontColor("#00FFFF");}],
    ["Color: Black", function(){setFontColor("#000000");}],
    ["Color: Blue", function(){setFontColor("#0000FF");}],
    ["Color: Brown", function(){setFontColor("#993300");}],
    ["Color: Dark Blue", function(){setFontColor("#000080");}],
    ["Color: Dark Cyan", function(){setFontColor("#008080");}],
    ["Color: Dark Gray", function(){setFontColor("#404040");}],
    ["Color: Dark Green", function(){setFontColor("#006400");}],
    ["Color: Dark Magenta", function(){setFontColor("#800080");}],
    ["Color: Dark Red", function(){setFontColor("#8B0000");}],
    ["Color: Dark Yellow", function(){setFontColor("#808000");}],
    ["Color: Gold", function(){setFontColor("#FFD700");}],
    ["Color: Gray", function(){setFontColor("#808080");}],
    ["Color: Green", function(){setFontColor("#008000");}],
    ["Color: Indigo", function(){setFontColor("#4B0082");}],
    ["Color: Lavender", function(){setFontColor("#E6E6FA");}],
    ["Color: Light Blue", function(){setFontColor("#ADD8E6");}],
    ["Color: Light Gray", function(){setFontColor("#C0C0C0");}],
    ["Color: Light Green", function(){setFontColor("#90EE90");}],
    ["Color: Lime", function(){setFontColor("#00FF00");}],
    ["Color: Magenta", function(){setFontColor("#FF00FF");}],
    ["Color: Maroon", function(){setFontColor("#800000");}],
    ["Color: Navy", function(){setFontColor("#000080");}],
    ["Color: Olive", function(){setFontColor("#808000");}],
    ["Color: Orange", function(){setFontColor("#FFA500");}],
    ["Color: Peach", function(){setFontColor("#FFDAB9");}],
    ["Color: Pink", function(){setFontColor("#FFC0CB");}],
    ["Color: Plum", function(){setFontColor("#DDA0DD");}],
    ["Color: Purple", function(){setFontColor("#800080");}],
    ["Color: Red", function(){setFontColor("#FF0000");}],
    ["Color: Rose", function(){setFontColor("#FF007F");}],
    ["Color: Silver", function(){setFontColor("#C0C0C0");}],
    ["Color: Sky Blue", function(){setFontColor("#87CEEB");}],
    ["Color: Tan", function(){setFontColor("#D2B48C");}],
    ["Color: Teal", function(){setFontColor("#008080");}],
    ["Color: Turquoise", function(){setFontColor("#40E0D0");}],
    ["Color: Violet", function(){setFontColor("#EE82EE");}],
    ["Color: White", function(){setFontColor("#FFFFFF");}],
    ["Color: Yellow", function(){setFontColor("#FFFF00");}],
    ["Columns: One", function(){ribbon("Layout","Columns > One");}],
    ["Columns: Three", function(){ribbon("Layout","Columns > Three");}],
    ["Columns: Two", function(){ribbon("Layout","Columns > Two");}],
    ["Combine Characters", function(){ribbon("Home","Combine Characters (Asian Layout)");}],
    ["Combine Documents", function(){ribbon("Review","Compare > Combine");}],
    ["Compare Documents", function(){ribbon("Review","Compare > Compare");}],
    ["Compress Pictures", function(){ribbon("Picture Format","Compress Pictures");}],
    ["Connector: Curved", function(){ribbon("Insert","Shapes > Lines > Curved Connector");}],
    ["Connector: Curved Arrow", function(){ribbon("Insert","Shapes > Lines > Curved Arrow Connector");}],
    ["Connector: Elbow", function(){ribbon("Insert","Shapes > Lines > Elbow Connector");}],
    ["Connector: Elbow Arrow", function(){ribbon("Insert","Shapes > Lines > Elbow Arrow Connector");}],
    ["Connector: Straight", function(){ribbon("Insert","Shapes > Lines > Straight Connector");}],
    ["Connector: Straight Arrow", function(){ribbon("Insert","Shapes > Lines > Straight Arrow Connector");}],
    ["Content Control Properties", function(){ribbon("Developer","Properties (select control first)");}],
    ["Content Control: Building Block Gallery", function(){ribbon("Developer","Building Block Gallery Content Control");}],
    ["Content Control: Check Box", function(){ribbon("Developer","Check Box Content Control");}],
    ["Content Control: Combo Box", function(){ribbon("Developer","Combo Box Content Control");}],
    ["Content Control: Date Picker", function(){ribbon("Developer","Date Picker Content Control");}],
    ["Content Control: Drop-Down List", function(){ribbon("Developer","Drop-Down List Content Control");}],
    ["Content Control: Legacy Check Box", function(){ribbon("Developer","Legacy Tools > Check Box Form Field");}],
    ["Content Control: Legacy Drop-Down", function(){ribbon("Developer","Legacy Tools > Drop-Down Form Field");}],
    ["Content Control: Legacy Text Field", function(){ribbon("Developer","Legacy Tools > Text Form Field");}],
    ["Content Control: Picture", function(){ribbon("Developer","Picture Content Control");}],
    ["Content Control: Plain Text", function(){ribbon("Developer","Plain Text Content Control");}],
    ["Content Control: Repeating Section", function(){ribbon("Developer","Repeating Section Content Control");}],
    ["Content Control: Rich Text", function(){ribbon("Developer","Rich Text Content Control");}],
    ["Continue Numbering", function(){ribbon("Home","Right-click numbered list > Continue Numbering");}],
    ["Convert Table to Text", function(){ribbon("Table Layout","Convert to Text");}],
    ["Convert Text to Table", function(){ribbon("Insert","Table > Convert Text to Table");}],
    ["Copilot: Draft with Copilot", function(){ribbon("Home","Copilot > Draft with Copilot");}],
    ["Copilot: Explain", function(){ribbon("Home","Copilot > Explain");}],
    ["Copilot: Rewrite", function(){ribbon("Home","Copilot > Rewrite");}],
    ["Copilot: Summarize", function(){ribbon("Home","Copilot > Summarize");}],
    ["Copilot: Undo Copilot Edit", function(){ribbon("Home","Copilot > Undo Copilot Edit");}],
    ["Copy", function(){shortcut("Ctrl+C");}],
    ["Cover Page", function(){ribbon("Insert","Cover Page");}],
    ["Create AutoText", function(){shortcut("Alt+F3");}],
    ["Create Subdocument", function(){ribbon("Outlining","Master Document > Create");}],
    ["Cross-reference", function(){ribbon("References","Cross-reference");}],
    ["Custom Margins", function(){ribbon("Layout","Margins > Custom Margins");}],
    ["Customize Keyboard", function(){ribbon("File","Options > Customize Ribbon > Customize...");}],
    ["Customize Quick Access Toolbar", function(){ribbon("File","Options > Quick Access Toolbar");}],
    ["Customize Ribbon", function(){ribbon("File","Options > Customize Ribbon");}],
    ["Cut", function(){shortcut("Ctrl+X");}],
    ["Date and Time", function(){ribbon("Insert","Date & Time");}],
    ["Decrease Indent", function(){setIndent("left",0);}],
    ["Decrease List Level", function(){shortcut("Shift+Tab in list");}],
    ["Define New Bullet", function(){ribbon("Home","Bullets > Define New Bullet");}],
    ["Define New List Style", function(){ribbon("Home","Multilevel List > Define New List Style");}],
    ["Define New Multilevel List", function(){ribbon("Home","Multilevel List > Define New Multilevel List");}],
    ["Define New Number Format", function(){ribbon("Home","Numbering > Define New Number Format");}],
    ["Delete", function(){shortcut("Delete key");}],
    ["Delete All Comments in Document", function(){ribbon("Review","Delete > Delete All Comments");}],
    ["Delete Cells", function(){ribbon("Table Layout","Delete > Delete Cells");}],
    ["Delete Columns", function(){ribbon("Table Layout","Delete > Delete Columns");}],
    ["Delete Comment", function(){ribbon("Review","Delete Comment");}],
    ["Delete Page", function(){ribbon("Select page content, then Delete");}],
    ["Delete Rows", function(){ribbon("Table Layout","Delete > Delete Rows");}],
    ["Delete Subdocument", function(){ribbon("Outlining","Master Document > Delete (Unlink)");}],
    ["Delete Table", function(){ribbon("Table Layout","Delete > Delete Table");}],
    ["Demote (Outline)", function(){ribbon("Outlining","Demote");}],
    ["Demote to Body Text", function(){ribbon("Outlining","Demote to Body Text");}],
    ["Design Mode", function(){ribbon("Developer","Design Mode");}],
    ["Different First Page (Header/Footer)", function(){ribbon("Header & Footer","Different First Page");}],
    ["Different Odd & Even Pages", function(){ribbon("Header & Footer","Different Odd & Even Pages");}],
    ["Distribute Columns Evenly", function(){ribbon("Table Layout","Distribute Columns");}],
    ["Distribute Rows Evenly", function(){ribbon("Table Layout","Distribute Rows");}],
    ["Document Inspector", function(){ribbon("File","Info > Check for Issues > Inspect Document");}],
    ["Document Map Toggle", function(){ribbon("View","Navigation Pane checkbox");}],
    ["Document Protection", function(){ribbon("Review","Restrict Editing");}],
    ["Don't Hyphenate", function(){ribbon("Layout","Hyphenation > None");}],
    ["Dot Accent", function(){ribbon("Home","Font dialog > Effects > Dot Accent");}],
    ["Double Strikethrough", function(){ribbon("Home","Font dialog > Effects > Double Strikethrough");}],
    ["Double Underline", function(){shortcut("Ctrl+Shift+D");}],
    ["Draft View", function(){ribbon("View","Draft");}],
    ["Draw Table", function(){ribbon("Insert","Table > Draw Table");}],
    ["Draw Text Box", function(){ribbon("Insert","Text Box > Draw Text Box");}],
    ["Drawing Canvas", function(){ribbon("Insert","Shapes > New Drawing Canvas");}],
    ["Drawing Pen", function(){ribbon("Draw","Pens > Drawing Pen");}],
    ["Drop Cap: Dropped", function(){ribbon("Insert","Drop Cap > Dropped");}],
    ["Drop Cap: In Margin", function(){ribbon("Insert","Drop Cap > In Margin");}],
    ["Drop Cap: None", function(){ribbon("Insert","Drop Cap > None");}],
    ["Edit Footer", function(){ribbon("Insert","Footer > Edit Footer");}],
    ["Edit Header", function(){ribbon("Insert","Header > Edit Header");}],
    ["Edit Points (Shape)", function(){ribbon("Shape Format","Edit Shape > Edit Points");}],
    ["Editing Restrictions", function(){ribbon("Review","Restrict Editing");}],
    ["Editor", function(){ribbon("Home","Editor");}],
    ["Editor: Clarity Refinements", function(){ribbon("Home","Editor > Clarity");}],
    ["Editor: Conciseness", function(){ribbon("Home","Editor > Conciseness");}],
    ["Editor: Formality", function(){ribbon("Home","Editor > Formality");}],
    ["Editor: Grammar Corrections", function(){ribbon("Home","Editor > Grammar");}],
    ["Editor: Inclusiveness", function(){ribbon("Home","Editor > Inclusiveness");}],
    ["Editor: Punctuation Conventions", function(){ribbon("Home","Editor > Punctuation");}],
    ["Editor: Similarity Checker", function(){ribbon("Home","Editor > Similarity");}],
    ["Editor: Spelling Corrections", function(){ribbon("Home","Editor > Spelling");}],
    ["Effects (Theme)", function(){ribbon("Design","Effects");}],
    ["Email as Attachment", function(){ribbon("File","Share > Email");}],
    ["Embed Fonts", function(){ribbon("File","Options > Save > Embed Fonts");}],
    ["Enclose Characters", function(){ribbon("Home","Enclose Characters (Asian Layout)");}],
    ["Encrypt with Password", function(){ribbon("File","Info > Protect Document > Encrypt with Password");}],
    ["Endnote: Insert", function(){shortcut("Ctrl+Alt+D");}],
    ["Envelopes", function(){ribbon("Mailings","Envelopes");}],
    ["Equation", function(){shortcut("Alt+=");}],
    ["Eraser (Ink)", function(){ribbon("Draw","Eraser");}],
    ["Eraser (Table)", function(){ribbon("Table Layout","Eraser");}],
    ["Even Page Section Break", function(){wordInsertBreak("EvenPage");}],
    ["Expand All Headings", function(){ribbon("View","Outline > Expand All");}],
    ["Expand Drawing Canvas", function(){ribbon("Drawing Canvas Format","Expand");}],
    ["Expand Heading", function(){ribbon("Home","Right-click heading > Expand/Collapse > Expand");}],
    ["Export to PDF/XPS", function(){ribbon("File","Export > Create PDF/XPS");}],
    ["Field", function(){ribbon("Insert","Quick Parts > Field");}],
    ["Field Codes: Toggle", function(){shortcut("Alt+F9");}],
    ["File: Close", function(){shortcut("Ctrl+W");}],
    ["File: Info", function(){ribbon("File","Info");}],
    ["File: New", function(){shortcut("Ctrl+N");}],
    ["File: Open", function(){shortcut("Ctrl+O");}],
    ["File: Options", function(){ribbon("File","Options");}],
    ["File: Print", function(){shortcut("Ctrl+P");}],
    ["File: Save As", function(){shortcut("F12 or Ctrl+Shift+S");}],
    ["Find", function(){shortcut("Ctrl+F");}],
    ["Find and Replace", function(){shortcut("Ctrl+H");}],
    ["Find Next", function(){shortcut("Ctrl+G or F5");}],
    ["First Line Indent", function(){setIndent("first",36);}],
    ["Flip Horizontal", function(){ribbon("Shape Format","Rotate > Flip Horizontal");}],
    ["Flip Vertical", function(){ribbon("Shape Format","Rotate > Flip Vertical");}],
    ["Focus Mode", function(){ribbon("View","Focus");}],
    ["Font Dialog", function(){shortcut("Ctrl+D");}],
    ["Font Size: 10", function(){setFontSize(10);}],
    ["Font Size: 10.5", function(){setFontSize(10.5);}],
    ["Font Size: 11", function(){setFontSize(11);}],
    ["Font Size: 12", function(){setFontSize(12);}],
    ["Font Size: 14", function(){setFontSize(14);}],
    ["Font Size: 16", function(){setFontSize(16);}],
    ["Font Size: 18", function(){setFontSize(18);}],
    ["Font Size: 20", function(){setFontSize(20);}],
    ["Font Size: 22", function(){setFontSize(22);}],
    ["Font Size: 24", function(){setFontSize(24);}],
    ["Font Size: 26", function(){setFontSize(26);}],
    ["Font Size: 28", function(){setFontSize(28);}],
    ["Font Size: 36", function(){setFontSize(36);}],
    ["Font Size: 48", function(){setFontSize(48);}],
    ["Font Size: 72", function(){setFontSize(72);}],
    ["Font Size: 8", function(){setFontSize(8);}],
    ["Font Size: 9", function(){setFontSize(9);}],
    ["Font: Arial", function(){setFontName("Arial");}],
    ["Font: Arial Black", function(){setFontName("Arial Black");}],
    ["Font: Calibri", function(){setFontName("Calibri");}],
    ["Font: Calibri Light", function(){setFontName("Calibri Light");}],
    ["Font: Cambria", function(){setFontName("Cambria");}],
    ["Font: Century Gothic", function(){setFontName("Century Gothic");}],
    ["Font: Comic Sans MS", function(){setFontName("Comic Sans MS");}],
    ["Font: Consolas", function(){setFontName("Consolas");}],
    ["Font: Constantia", function(){setFontName("Constantia");}],
    ["Font: Corbel", function(){setFontName("Corbel");}],
    ["Font: Courier New", function(){setFontName("Courier New");}],
    ["Font: Franklin Gothic", function(){setFontName("Franklin Gothic Medium");}],
    ["Font: Garamond", function(){setFontName("Garamond");}],
    ["Font: Georgia", function(){setFontName("Georgia");}],
    ["Font: Impact", function(){setFontName("Impact");}],
    ["Font: Lucida Console", function(){setFontName("Lucida Console");}],
    ["Font: Lucida Sans", function(){setFontName("Lucida Sans");}],
    ["Font: Palatino Linotype", function(){setFontName("Palatino Linotype");}],
    ["Font: Segoe UI", function(){setFontName("Segoe UI");}],
    ["Font: Tahoma", function(){setFontName("Tahoma");}],
    ["Font: Times New Roman", function(){setFontName("Times New Roman");}],
    ["Font: Trebuchet MS", function(){setFontName("Trebuchet MS");}],
    ["Font: Verdana", function(){setFontName("Verdana");}],
    ["Footer", function(){ribbon("Insert","Footer");}],
    ["Footnote: Insert", function(){shortcut("Ctrl+Alt+F");}],
    ["Format Painter", function(){shortcut("Ctrl+Shift+C to copy, Ctrl+Shift+V to paste format");}],
    ["Formatting Marks (Show/Hide)", function(){shortcut("Ctrl+Shift+8 or Ctrl+*");}],
    ["Frame Format", function(){ribbon("Format","Frame (legacy or right-click frame)");}],
    ["Full Screen Reading", function(){ribbon("View","Read Mode");}],
    ["Go Back", function(){shortcut("Alt+Left Arrow");}],
    ["Go Forward", function(){shortcut("Alt+Right Arrow");}],
    ["Go To", function(){shortcut("Ctrl+G or F5");}],
    ["Go to Bookmark", function(){shortcut("Ctrl+Shift+F5");}],
    ["Go to Footer", function(){ribbon("Insert","Footer > Edit Footer");}],
    ["Go to Header", function(){ribbon("Insert","Header > Edit Header");}],
    ["Go to Next Comment", function(){ribbon("Review","Next Comment");}],
    ["Go to Next Section", function(){ribbon("Navigate: Ctrl+G > Section");}],
    ["Go to Previous Comment", function(){ribbon("Review","Previous Comment");}],
    ["Greeting Line (Mail Merge)", function(){ribbon("Mailings","Greeting Line");}],
    ["Gridlines (View)", function(){ribbon("View","Gridlines");}],
    ["Group Objects", function(){ribbon("Shape Format","Group > Group");}],
    ["Grow Font", function(){shortcut("Ctrl+Shift+>");}],
    ["Hanging Indent", function(){setIndent("first",-36);}],
    ["Header", function(){ribbon("Insert","Header");}],
    ["Heading Rows Repeat", function(){ribbon("Table Layout","Repeat Header Rows");}],
    ["Help", function(){shortcut("F1");}],
    ["Hidden Text", function(){ribbon("Home","Font dialog > Effects > Hidden");}],
    ["Highlight: Blue", function(){setHighlight("#0000FF");}],
    ["Highlight: Bright Green", function(){setHighlight("#00FF00");}],
    ["Highlight: Cyan", function(){setHighlight("#00FFFF");}],
    ["Highlight: Dark Blue", function(){setHighlight("#000080");}],
    ["Highlight: Dark Red", function(){setHighlight("#800000");}],
    ["Highlight: Dark Yellow", function(){setHighlight("#808000");}],
    ["Highlight: Gray 25%", function(){setHighlight("#C0C0C0");}],
    ["Highlight: Gray 50%", function(){setHighlight("#808080");}],
    ["Highlight: Green", function(){setHighlight("#008000");}],
    ["Highlight: Pink", function(){setHighlight("#FF00FF");}],
    ["Highlight: Red", function(){setHighlight("#FF0000");}],
    ["Highlight: Remove", function(){setHighlight(null);}],
    ["Highlight: Teal", function(){setHighlight("#008080");}],
    ["Highlight: Yellow", function(){setHighlight("#FFFF00");}],
    ["Highlighter Pen", function(){ribbon("Draw","Pens > Highlighter");}],
    ["Horizontal Line", function(){wordInsertHtml('<hr style="border:1px solid #999;width:100%">',"Horizontal line inserted.");}],
    ["Hyperlink: Insert", function(){shortcut("Ctrl+K");}],
    ["Hyphenation: Automatic", function(){ribbon("Layout","Hyphenation > Automatic");}],
    ["Hyphenation: Manual", function(){ribbon("Layout","Hyphenation > Manual");}],
    ["Hyphenation: None", function(){ribbon("Layout","Hyphenation > None");}],
    ["Icons", function(){ribbon("Insert","Icons");}],
    ["Immersive Reader", function(){ribbon("View","Immersive Reader");}],
    ["Import Subdocument", function(){ribbon("Outlining","Master Document > Insert");}],
    ["Increase Indent", function(){setIndent("left",36);}],
    ["Increase List Level", function(){shortcut("Tab in list");}],
    ["Index: Insert", function(){ribbon("References","Insert Index");}],
    ["Index: Mark Entry", function(){shortcut("Alt+Shift+X");}],
    ["Index: Update", function(){ribbon("References","Update Index");}],
    ["Ink Annotations: Toggle", function(){ribbon("Review","Ink > Start/Stop Inking");}],
    ["Ink Replay", function(){ribbon("Draw","Ink Replay");}],
    ["Ink to Math", function(){ribbon("Draw","Ink to Math");}],
    ["Ink to Shape", function(){ribbon("Draw","Ink to Shape");}],
    ["Ink to Text", function(){ribbon("Draw","Ink to Text");}],
    ["Insert Address Block (Mail Merge)", function(){ribbon("Mailings","Address Block");}],
    ["Insert Caption", function(){ribbon("References","Insert Caption");}],
    ["Insert Cells", function(){ribbon("Table Layout","Insert > Insert Cells");}],
    ["Insert Citation", function(){ribbon("References","Insert Citation");}],
    ["Insert Columns to the Left", function(){ribbon("Table Layout","Insert Left");}],
    ["Insert Columns to the Right", function(){ribbon("Table Layout","Insert Right");}],
    ["Insert Comment", function(){shortcut("Ctrl+Alt+M");}],
    ["Insert Date", function(){ribbon("Insert","Date & Time");}],
    ["Insert Endnote", function(){shortcut("Ctrl+Alt+D");}],
    ["Insert Equation", function(){shortcut("Alt+=");}],
    ["Insert Field", function(){ribbon("Insert","Quick Parts > Field");}],
    ["Insert File (Text from File)", function(){ribbon("Insert","Object > Text from File");}],
    ["Insert Footnote", function(){shortcut("Ctrl+Alt+F");}],
    ["Insert Frame", function(){ribbon("Developer","Legacy Tools > Frame");}],
    ["Insert Merge Field", function(){ribbon("Mailings","Insert Merge Field");}],
    ["Insert Online Pictures", function(){ribbon("Insert","Online Pictures");}],
    ["Insert Online Video", function(){ribbon("Insert","Online Video");}],
    ["Insert Page Number", function(){ribbon("Insert","Page Number");}],
    ["Insert Picture from File", function(){ribbon("Insert","Pictures > This Device");}],
    ["Insert Rows Above", function(){ribbon("Table Layout","Insert Above");}],
    ["Insert Rows Below", function(){ribbon("Table Layout","Insert Below");}],
    ["Insert Signature Line", function(){ribbon("Insert","Signature Line");}],
    ["Insert Symbol", function(){ribbon("Insert","Symbol > More Symbols");}],
    ["Insert Table of Authorities", function(){ribbon("References","Insert Table of Authorities");}],
    ["Insert Table of Contents", function(){ribbon("References","Table of Contents");}],
    ["Insert Table of Figures", function(){ribbon("References","Insert Table of Figures");}],
    ["Insert Text Box", function(){ribbon("Insert","Text Box");}],
    ["Insert Time", function(){ribbon("Insert","Date & Time (with time format)");}],
    ["Insert WordArt", function(){ribbon("Insert","WordArt");}],
    ["Italic", function(){toggleFont("italic");}],
    ["Join to Previous List", function(){ribbon("Home","Right-click numbered list > Join to Previous List");}],
    ["Justify", function(){setAlignment("Justified");}],
    ["Keep Lines Together", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Keep lines together");}],
    ["Keep with Next", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Keep with next");}],
    ["Labels", function(){ribbon("Mailings","Labels");}],
    ["Landscape Orientation", function(){ribbon("Layout","Orientation > Landscape");}],
    ["Language: Set Proofing", function(){ribbon("Review","Language > Set Proofing Language");}],
    ["Language: Translate Document", function(){ribbon("Review","Translate > Translate Document");}],
    ["Language: Translate Selection", function(){ribbon("Review","Translate > Translate Selection");}],
    ["Lasso Select (Ink)", function(){ribbon("Draw","Lasso Select");}],
    ["Last Column (Table Style)", function(){ribbon("Table Design","Last Column");}],
    ["Left Indent", function(){setIndent("left",36);}],
    ["Left Tab Stop", function(){ribbon("Home","Paragraph dialog > Tabs");}],
    ["Line and Page Breaks", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks");}],
    ["Line Break", function(){shortcut("Shift+Enter");}],
    ["Line Numbers", function(){ribbon("Layout","Line Numbers");}],
    ["Line Spacing: 1.0", function(){setLineSpacing(12);}],
    ["Line Spacing: 1.15", function(){setLineSpacing(13.8);}],
    ["Line Spacing: 1.5", function(){setLineSpacing(18);}],
    ["Line Spacing: 2.0", function(){setLineSpacing(24);}],
    ["Line Spacing: 2.5", function(){setLineSpacing(30);}],
    ["Line Spacing: 3.0", function(){setLineSpacing(36);}],
    ["Link to Previous (Header/Footer)", function(){ribbon("Header & Footer","Link to Previous");}],
    ["Lock Subdocument", function(){ribbon("Outlining","Master Document > Lock Document");}],
    ["Lock Tracking", function(){ribbon("Review","Track Changes > Lock Tracking");}],
    ["Macros: Record", function(){ribbon("View","Macros > Record Macro");}],
    ["Macros: View", function(){shortcut("Alt+F8");}],
    ["Mail Merge: Check for Errors", function(){ribbon("Mailings","Finish & Merge > Check for Errors");}],
    ["Mail Merge: Directory", function(){ribbon("Mailings","Start Mail Merge > Directory");}],
    ["Mail Merge: Edit Individual Documents", function(){ribbon("Mailings","Finish & Merge > Edit Individual Documents");}],
    ["Mail Merge: Edit Recipient List", function(){ribbon("Mailings","Edit Recipient List");}],
    ["Mail Merge: Email Messages", function(){ribbon("Mailings","Start Mail Merge > E-Mail Messages");}],
    ["Mail Merge: Envelopes", function(){ribbon("Mailings","Start Mail Merge > Envelopes");}],
    ["Mail Merge: Find Recipient", function(){ribbon("Mailings","Find Recipient");}],
    ["Mail Merge: Finish & Merge", function(){ribbon("Mailings","Finish & Merge");}],
    ["Mail Merge: First Record", function(){ribbon("Mailings","Preview Results > First Record");}],
    ["Mail Merge: Go to Record", function(){ribbon("Mailings","Preview Results > Go to Record");}],
    ["Mail Merge: Highlight Merge Fields", function(){ribbon("Mailings","Highlight Merge Fields");}],
    ["Mail Merge: Labels", function(){ribbon("Mailings","Start Mail Merge > Labels");}],
    ["Mail Merge: Last Record", function(){ribbon("Mailings","Preview Results > Last Record");}],
    ["Mail Merge: Letters", function(){ribbon("Mailings","Start Mail Merge > Letters");}],
    ["Mail Merge: Match Fields", function(){ribbon("Mailings","Match Fields");}],
    ["Mail Merge: Next Record", function(){ribbon("Mailings","Preview Results > Next Record");}],
    ["Mail Merge: Normal Word Document", function(){ribbon("Mailings","Start Mail Merge > Normal Word Document");}],
    ["Mail Merge: Preview Results", function(){ribbon("Mailings","Preview Results");}],
    ["Mail Merge: Previous Record", function(){ribbon("Mailings","Preview Results > Previous Record");}],
    ["Mail Merge: Print Documents", function(){ribbon("Mailings","Finish & Merge > Print Documents");}],
    ["Mail Merge: Rules: Ask", function(){ribbon("Mailings","Rules > Ask");}],
    ["Mail Merge: Rules: Fill-in", function(){ribbon("Mailings","Rules > Fill-in");}],
    ["Mail Merge: Rules: If Then Else", function(){ribbon("Mailings","Rules > If...Then...Else");}],
    ["Mail Merge: Rules: Merge Record", function(){ribbon("Mailings","Rules > Merge Record #");}],
    ["Mail Merge: Rules: Merge Sequence", function(){ribbon("Mailings","Rules > Merge Sequence #");}],
    ["Mail Merge: Rules: Next Record", function(){ribbon("Mailings","Rules > Next Record");}],
    ["Mail Merge: Rules: Next Record If", function(){ribbon("Mailings","Rules > Next Record If");}],
    ["Mail Merge: Rules: Set Bookmark", function(){ribbon("Mailings","Rules > Set Bookmark");}],
    ["Mail Merge: Rules: Skip Record If", function(){ribbon("Mailings","Rules > Skip Record If");}],
    ["Mail Merge: Select Recipients", function(){ribbon("Mailings","Select Recipients");}],
    ["Mail Merge: Send Email Messages", function(){ribbon("Mailings","Finish & Merge > Send Email Messages");}],
    ["Mail Merge: Start", function(){ribbon("Mailings","Start Mail Merge");}],
    ["Mail Merge: Step-by-Step Wizard", function(){ribbon("Mailings","Start Mail Merge > Step-by-Step Wizard");}],
    ["Manage Add-ins", function(){ribbon("Insert","My Add-ins > Manage My Add-ins");}],
    ["Manage Sources", function(){ribbon("References","Manage Sources");}],
    ["Manage Styles", function(){ribbon("Home","Styles launcher > Manage Styles");}],
    ["Manual Hyphenation", function(){ribbon("Layout","Hyphenation > Manual");}],
    ["Margins: Mirrored", function(){ribbon("Layout","Margins > Mirrored");}],
    ["Margins: Moderate", function(){ribbon("Layout","Margins > Moderate");}],
    ["Margins: Narrow", function(){ribbon("Layout","Margins > Narrow");}],
    ["Margins: Normal", function(){ribbon("Layout","Margins > Normal");}],
    ["Margins: Wide", function(){ribbon("Layout","Margins > Wide");}],
    ["Mark Citation", function(){ribbon("References","Mark Citation");}],
    ["Mark Index Entry", function(){shortcut("Alt+Shift+X");}],
    ["Mark Table of Contents Entry", function(){ribbon("References","Add Text");}],
    ["Merge Cells", function(){ribbon("Table Layout","Merge Cells");}],
    ["Merge Formatting (Paste)", function(){shortcut("Ctrl+Shift+V (then choose)");}],
    ["Merge Subdocuments", function(){ribbon("Outlining","Master Document > Merge");}],
    ["Modify Style", function(){ribbon("Home","Styles > right-click style > Modify");}],
    ["Move Down (Outline)", function(){ribbon("Outlining","Move Down");}],
    ["Move Up (Outline)", function(){ribbon("Outlining","Move Up");}],
    ["Multilevel List", function(){ribbon("Home","Multilevel List");}],
    ["Navigation Pane", function(){shortcut("Ctrl+F (opens Navigation)");}],
    ["Navigation Pane: Headings Tab", function(){ribbon("View","Navigation Pane > Headings tab");}],
    ["Navigation Pane: Pages Tab", function(){ribbon("View","Navigation Pane > Pages tab");}],
    ["Navigation Pane: Results Tab", function(){ribbon("View","Navigation Pane > Results tab");}],
    ["New Blank Document", function(){shortcut("Ctrl+N");}],
    ["New Comment", function(){shortcut("Ctrl+Alt+M");}],
    ["New Folder", function(){ribbon("File","Save As > New Folder");}],
    ["New from Template", function(){ribbon("File","New");}],
    ["New Window", function(){ribbon("View","New Window");}],
    ["Next Change", function(){ribbon("Review","Next Change");}],
    ["Next Comment", function(){ribbon("Review","Next Comment");}],
    ["Next Footnote", function(){ribbon("References","Next Footnote");}],
    ["Next Page Section Break", function(){wordInsertBreak("SectionNext");}],
    ["Numbering", function(){ribbon("Home","Numbering");}],
    ["Object: Insert", function(){ribbon("Insert","Object");}],
    ["Odd Page Section Break", function(){wordInsertBreak("OddPage");}],
    ["Open", function(){shortcut("Ctrl+O");}],
    ["Open Hyperlink", function(){shortcut("Ctrl+Click on hyperlink");}],
    ["Open in Browser", function(){ribbon("File","Info > Open in Browser");}],
    ["Orientation: Landscape", function(){ribbon("Layout","Orientation > Landscape");}],
    ["Orientation: Portrait", function(){ribbon("Layout","Orientation > Portrait");}],
    ["Outline View", function(){ribbon("View","Outline");}],
    ["Outline: Show All Levels", function(){ribbon("Outlining","Show Level > All Levels");}],
    ["Outline: Show First Line Only", function(){ribbon("Outlining","Show First Line Only checkbox");}],
    ["Outline: Show Level 1", function(){ribbon("Outlining","Show Level > Level 1");}],
    ["Outline: Show Level 2", function(){ribbon("Outlining","Show Level > Level 2");}],
    ["Outline: Show Level 3", function(){ribbon("Outlining","Show Level > Level 3");}],
    ["Outline: Show Level 4", function(){ribbon("Outlining","Show Level > Level 4");}],
    ["Outline: Show Level 5", function(){ribbon("Outlining","Show Level > Level 5");}],
    ["Outline: Show Level 6", function(){ribbon("Outlining","Show Level > Level 6");}],
    ["Outline: Show Level 7", function(){ribbon("Outlining","Show Level > Level 7");}],
    ["Outline: Show Level 8", function(){ribbon("Outlining","Show Level > Level 8");}],
    ["Outline: Show Level 9", function(){ribbon("Outlining","Show Level > Level 9");}],
    ["Outline: Show Text Formatting", function(){ribbon("Outlining","Show Text Formatting checkbox");}],
    ["Page Background: Color", function(){ribbon("Design","Page Color");}],
    ["Page Background: Page Borders", function(){ribbon("Design","Page Borders");}],
    ["Page Background: Watermark", function(){ribbon("Design","Watermark");}],
    ["Page Break", function(){wordInsertBreak("Page");}],
    ["Page Break Before", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Page break before");}],
    ["Page Color", function(){ribbon("Design","Page Color");}],
    ["Page Down", function(){shortcut("Page Down key");}],
    ["Page Layout View", function(){ribbon("View","Print Layout");}],
    ["Page Number: Bottom of Page", function(){ribbon("Insert","Page Number > Bottom of Page");}],
    ["Page Number: Current Position", function(){ribbon("Insert","Page Number > Current Position");}],
    ["Page Number: Format", function(){ribbon("Insert","Page Number > Format Page Numbers");}],
    ["Page Number: Page Margins", function(){ribbon("Insert","Page Number > Page Margins");}],
    ["Page Number: Remove", function(){ribbon("Insert","Page Number > Remove Page Numbers");}],
    ["Page Number: Top of Page", function(){ribbon("Insert","Page Number > Top of Page");}],
    ["Page Setup Dialog", function(){ribbon("Layout","Page Setup launcher (bottom-right arrow)");}],
    ["Page Up", function(){shortcut("Page Up key");}],
    ["Paragraph Dialog", function(){ribbon("Home","Paragraph launcher (bottom-right arrow)");}],
    ["Paragraph Marks (Show/Hide)", function(){shortcut("Ctrl+Shift+8");}],
    ["Paragraph Shading", function(){ribbon("Home","Shading (paint bucket)");}],
    ["Paste", function(){shortcut("Ctrl+V");}],
    ["Paste as Hyperlink", function(){ribbon("Home","Paste > Paste Special > Paste as Hyperlink");}],
    ["Paste as Picture", function(){ribbon("Home","Paste > Paste Special > Picture");}],
    ["Paste Special", function(){shortcut("Ctrl+Alt+V");}],
    ["Paste: Keep Source Formatting", function(){shortcut("Ctrl+V then Ctrl > K");}],
    ["Paste: Keep Text Only", function(){shortcut("Ctrl+V then Ctrl > T");}],
    ["Paste: Merge Formatting", function(){shortcut("Ctrl+V then Ctrl > M");}],
    ["Phonetic Guide", function(){ribbon("Home","Phonetic Guide (Asian Layout)");}],
    ["Picture: Border", function(){ribbon("Picture Format","Picture Border");}],
    ["Picture: Compress", function(){ribbon("Picture Format","Compress Pictures");}],
    ["Picture: Crop", function(){ribbon("Picture Format","Crop");}],
    ["Picture: Effects", function(){ribbon("Picture Format","Picture Effects");}],
    ["Picture: Insert", function(){ribbon("Insert","Pictures");}],
    ["Picture: Layout", function(){ribbon("Picture Format","Picture Layout");}],
    ["Picture: Reset", function(){ribbon("Picture Format","Reset Picture");}],
    ["Portrait Orientation", function(){ribbon("Layout","Orientation > Portrait");}],
    ["Position Object", function(){ribbon("Shape Format","Position");}],
    ["Present Online", function(){ribbon("File","Share > Present Online");}],
    ["Previous Change", function(){ribbon("Review","Previous Change");}],
    ["Previous Comment", function(){ribbon("Review","Previous Comment");}],
    ["Print", function(){shortcut("Ctrl+P");}],
    ["Print Layout View", function(){ribbon("View","Print Layout");}],
    ["Print Preview", function(){shortcut("Ctrl+F2");}],
    ["Promote (Outline)", function(){ribbon("Outlining","Promote");}],
    ["Promote to Heading 1", function(){ribbon("Outlining","Promote to Heading 1");}],
    ["Protect Document", function(){ribbon("File","Info > Protect Document");}],
    ["Quick Parts", function(){ribbon("Insert","Quick Parts");}],
    ["Quick Print", function(){shortcut("Ctrl+P then Enter");}],
    ["Quick Styles", function(){ribbon("Home","Styles gallery");}],
    ["Quick Tables", function(){ribbon("Insert","Table > Quick Tables");}],
    ["Read Aloud", function(){ribbon("Review","Read Aloud");}],
    ["Read Mode", function(){ribbon("View","Read Mode");}],
    ["Redo", function(){shortcut("Ctrl+Y");}],
    ["References: Bibliography", function(){ribbon("References","Bibliography");}],
    ["References: Manage Sources", function(){ribbon("References","Manage Sources");}],
    ["References: Style (APA, MLA, etc.)", function(){ribbon("References","Style dropdown");}],
    ["Reject All Changes in Document", function(){ribbon("Review","Reject > Reject All Changes");}],
    ["Reject and Move to Next", function(){ribbon("Review","Reject");}],
    ["Remove Content Control", function(){ribbon("Developer","Right-click control > Remove Content Control");}],
    ["Remove Hyperlink", function(){ribbon("Right-click hyperlink > Remove Hyperlink");}],
    ["Remove Page Break", function(){showToast("Show formatting marks (Ctrl+Shift+8), select break, press Delete.");}],
    ["Remove Space After Paragraph", function(){setSpaceAfter(0);}],
    ["Remove Space Before Paragraph", function(){setSpaceBefore(0);}],
    ["Remove Table of Contents", function(){ribbon("References","Table of Contents > Remove Table of Contents");}],
    ["Remove Watermark", function(){ribbon("Design","Watermark > Remove Watermark");}],
    ["Repeat", function(){shortcut("Ctrl+Y or F4");}],
    ["Repeat Header Rows", function(){ribbon("Table Layout","Repeat Header Rows");}],
    ["Replace", function(){shortcut("Ctrl+H");}],
    ["Research", function(){ribbon("Review","Research / Smart Lookup");}],
    ["Reset Character Formatting", function(){shortcut("Ctrl+Space");}],
    ["Reset Paragraph Formatting", function(){shortcut("Ctrl+Q");}],
    ["Reset Picture", function(){ribbon("Picture Format","Reset Picture");}],
    ["Restart Numbering", function(){ribbon("Home","right-click number > Restart at 1");}],
    ["Restrict Editing", function(){ribbon("Review","Restrict Editing");}],
    ["Review in Editor", function(){ribbon("Home","Editor > Review in Editor pane");}],
    ["Reviewing Pane: Horizontal", function(){ribbon("Review","Reviewing Pane > Horizontal");}],
    ["Reviewing Pane: Vertical", function(){ribbon("Review","Reviewing Pane > Vertical");}],
    ["Right Indent", function(){setIndent("right",36);}],
    ["Right Tab Stop", function(){ribbon("Home","Paragraph dialog > Tabs");}],
    ["Rotate Left 90", function(){ribbon("Shape Format","Rotate > Rotate Left 90");}],
    ["Rotate Right 90", function(){ribbon("Shape Format","Rotate > Rotate Right 90");}],
    ["Ruler (Draw)", function(){ribbon("Draw","Ruler");}],
    ["Ruler (Show/Hide)", function(){ribbon("View","Ruler");}],
    ["Save All", function(){shortcut("Not available via API. Save each document individually.");}],
    ["Save As", function(){shortcut("F12 or Ctrl+Shift+S");}],
    ["Save As PDF", function(){ribbon("File","Export > Create PDF/XPS");}],
    ["Save as Picture", function(){ribbon("Right-click object","Save as Picture");}],
    ["Save As Template", function(){ribbon("File","Save As > Word Template (*.dotx)");}],
    ["Screen Clipping", function(){ribbon("Insert","Screenshot > Screen Clipping");}],
    ["Screenshot", function(){ribbon("Insert","Screenshot");}],
    ["Search Document", function(){shortcut("Ctrl+F");}],
    ["Section Break: Continuous", function(){wordInsertBreak("SectionContinuous");}],
    ["Section Break: Even Page", function(){wordInsertBreak("EvenPage");}],
    ["Section Break: Next Page", function(){wordInsertBreak("SectionNext");}],
    ["Section Break: Odd Page", function(){wordInsertBreak("OddPage");}],
    ["Select Cell", function(){ribbon("Table Layout","Select > Select Cell");}],
    ["Select Column", function(){ribbon("Table Layout","Select > Select Column");}],
    ["Select Objects", function(){ribbon("Home","Select > Select Objects");}],
    ["Select Row", function(){ribbon("Table Layout","Select > Select Row");}],
    ["Select Table", function(){ribbon("Table Layout","Select > Select Table");}],
    ["Selection Pane", function(){ribbon("Shape Format","Selection Pane");}],
    ["Send as Attachment", function(){ribbon("File","Share > Email");}],
    ["Send as PDF", function(){ribbon("File","Share > Email > Send as PDF");}],
    ["Send Backward", function(){ribbon("Shape Format","Send Backward");}],
    ["Send Behind Text", function(){ribbon("Shape Format","Send Backward > Send Behind Text");}],
    ["Send to Back", function(){ribbon("Shape Format","Send to Back");}],
    ["Separate List", function(){ribbon("Home","Right-click numbered list > Separate List");}],
    ["Set Numbering Value", function(){ribbon("Home","Right-click numbered list > Set Numbering Value");}],
    ["Shading", function(){ribbon("Home","Shading (paint bucket icon)");}],
    ["Shape Effects", function(){ribbon("Shape Format","Shape Effects");}],
    ["Shape Fill", function(){ribbon("Shape Format","Shape Fill");}],
    ["Shape Outline", function(){ribbon("Shape Format","Shape Outline");}],
    ["Shapes: Insert", function(){ribbon("Insert","Shapes");}],
    ["Share", function(){ribbon("File","Share");}],
    ["Show All Formatting Marks", function(){shortcut("Ctrl+Shift+8 or Ctrl+*");}],
    ["Show Comments", function(){ribbon("Review","Show Comments");}],
    ["Show Markup", function(){ribbon("Review","Show Markup");}],
    ["Show Subdocuments", function(){ribbon("Outlining","Master Document > Show Subdocuments");}],
    ["Show/Hide Paragraph Marks", function(){shortcut("Ctrl+Shift+8");}],
    ["Shrink Font", function(){shortcut("Ctrl+Shift+<");}],
    ["Shrink One Page", function(){ribbon("Print Preview","Shrink One Page");}],
    ["Sign Document (Digital Signature)", function(){ribbon("File","Info > Protect Document > Add a Digital Signature");}],
    ["Signature Details", function(){ribbon("Right-click signature","Signature Details");}],
    ["Signature: Remove", function(){ribbon("Right-click signature","Remove Signature");}],
    ["Simple Markup", function(){ribbon("Review","Display for Review > Simple Markup");}],
    ["Size: A4", function(){ribbon("Layout","Size > A4");}],
    ["Size: A5", function(){ribbon("Layout","Size > A5");}],
    ["Size: B5", function(){ribbon("Layout","Size > B5");}],
    ["Size: Executive", function(){ribbon("Layout","Size > Executive");}],
    ["Size: Legal (8.5x14)", function(){ribbon("Layout","Size > Legal");}],
    ["Size: Letter (8.5x11)", function(){ribbon("Layout","Size > Letter");}],
    ["Small Caps", function(){shortcut("Ctrl+Shift+K");}],
    ["SmartArt", function(){ribbon("Insert","SmartArt");}],
    ["Snap to Grid", function(){ribbon("Shape Format","Align > Grid Settings > Snap to Grid");}],
    ["Sort", function(){ribbon("Home","Sort (A-Z icon)");}],
    ["Sort Ascending", function(){ribbon("Home","Sort > Ascending");}],
    ["Sort Descending", function(){ribbon("Home","Sort > Descending");}],
    ["Space After: 0 pt", function(){setSpaceAfter(0);}],
    ["Space After: 10 pt", function(){setSpaceAfter(10);}],
    ["Space After: 12 pt", function(){setSpaceAfter(12);}],
    ["Space After: 24 pt", function(){setSpaceAfter(24);}],
    ["Space After: 6 pt", function(){setSpaceAfter(6);}],
    ["Space After: 8 pt", function(){setSpaceAfter(8);}],
    ["Space Before: 0 pt", function(){setSpaceBefore(0);}],
    ["Space Before: 10 pt", function(){setSpaceBefore(10);}],
    ["Space Before: 12 pt", function(){setSpaceBefore(12);}],
    ["Space Before: 24 pt", function(){setSpaceBefore(24);}],
    ["Space Before: 6 pt", function(){setSpaceBefore(6);}],
    ["Space Before: 8 pt", function(){setSpaceBefore(8);}],
    ["Spacing: 1.5", function(){setLineSpacing(18);}],
    ["Spacing: Double", function(){setLineSpacing(24);}],
    ["Spacing: Single", function(){setLineSpacing(12);}],
    ["Spelling and Grammar", function(){shortcut("F7");}],
    ["Split Cells", function(){ribbon("Table Layout","Split Cells");}],
    ["Split Subdocument", function(){ribbon("Outlining","Master Document > Split");}],
    ["Split Table", function(){ribbon("Table Layout","Split Table");}],
    ["Split Window", function(){ribbon("View","Split");}],
    ["Start Mail Merge", function(){ribbon("Mailings","Start Mail Merge");}],
    ["Start Tracking Changes", function(){shortcut("Ctrl+Shift+E");}],
    ["Strikethrough", function(){toggleFont("strikethrough");}],
    ["Switch Windows", function(){ribbon("View","Switch Windows");}],
    ["Symbol: Insert", function(){ribbon("Insert","Symbol > More Symbols");}],
    ["Synonyms / Thesaurus", function(){shortcut("Shift+F7");}],
    ["Tab Character", function(){shortcut("Tab key");}],
    ["Table Design: Banded Columns", function(){ribbon("Table Design","Banded Columns");}],
    ["Table Design: Banded Rows", function(){ribbon("Table Design","Banded Rows");}],
    ["Table Design: First Column", function(){ribbon("Table Design","First Column");}],
    ["Table Design: Header Row", function(){ribbon("Table Design","Header Row");}],
    ["Table Design: Last Column", function(){ribbon("Table Design","Last Column");}],
    ["Table Design: Total Row", function(){ribbon("Table Design","Total Row");}],
    ["Table Gridlines (Show/Hide)", function(){ribbon("Table Layout","View Gridlines");}],
    ["Table of Authorities: Insert", function(){ribbon("References","Insert Table of Authorities");}],
    ["Table of Authorities: Mark Citation", function(){ribbon("References","Mark Citation");}],
    ["Table of Contents: Insert", function(){ribbon("References","Table of Contents");}],
    ["Table of Contents: Remove", function(){ribbon("References","Table of Contents > Remove");}],
    ["Table of Contents: Update", function(){ribbon("References","Update Table");}],
    ["Table of Figures: Insert", function(){ribbon("References","Insert Table of Figures");}],
    ["Table of Figures: Update", function(){ribbon("References","Update Table");}],
    ["Table Properties", function(){ribbon("Table Layout","Properties");}],
    ["Table Select", function(){ribbon("Table Layout","Select");}],
    ["Tabs Dialog", function(){ribbon("Home","Paragraph dialog > Tabs");}],
    ["Templates", function(){ribbon("File","New > Personal Templates");}],
    ["Text Box: Draw", function(){ribbon("Insert","Text Box > Draw Text Box");}],
    ["Text Box: Simple", function(){ribbon("Insert","Text Box > Simple Text Box");}],
    ["Text Direction", function(){ribbon("Table Layout","Text Direction");}],
    ["Text Effects", function(){ribbon("Home","Text Effects and Typography");}],
    ["Text Highlight Color", function(){ribbon("Home","Text Highlight Color");}],
    ["Text Predictions: Accept", function(){shortcut("Tab (when prediction is shown)");}],
    ["Text Predictions: Toggle", function(){ribbon("File","Options > Advanced > Editor > Text Predictions");}],
    ["Text Wrapping: Behind Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > Behind Text");}],
    ["Text Wrapping: In Front of Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > In Front of Text");}],
    ["Text Wrapping: In Line with Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > In Line with Text");}],
    ["Text Wrapping: Square", function(){ribbon("Shape Format/Picture Format","Wrap Text > Square");}],
    ["Text Wrapping: Through", function(){ribbon("Shape Format/Picture Format","Wrap Text > Through");}],
    ["Text Wrapping: Tight", function(){ribbon("Shape Format/Picture Format","Wrap Text > Tight");}],
    ["Text Wrapping: Top and Bottom", function(){ribbon("Shape Format/Picture Format","Wrap Text > Top and Bottom");}],
    ["Theme Colors", function(){ribbon("Design","Colors");}],
    ["Theme Effects", function(){ribbon("Design","Effects");}],
    ["Theme Fonts", function(){ribbon("Design","Fonts");}],
    ["Themes", function(){ribbon("Design","Themes");}],
    ["Thesaurus", function(){shortcut("Shift+F7");}],
    ["Toggle Field Codes", function(){shortcut("Alt+F9");}],
    ["Toggle Full Screen", function(){shortcut("Alt+V, U (or View > Full Screen)");}],
    ["Track Changes", function(){shortcut("Ctrl+Shift+E");}],
    ["Track Changes Options", function(){ribbon("Review","Track Changes > Change Tracking Options");}],
    ["Translate Document", function(){ribbon("Review","Translate > Translate Document");}],
    ["Translate Selection", function(){ribbon("Review","Translate > Translate Selection");}],
    ["Translate: Chinese (Simplified to Traditional)", function(){ribbon("Review","Translate > Chinese Simplified to Traditional");}],
    ["Translate: Chinese (Traditional to Simplified)", function(){ribbon("Review","Translate > Chinese Traditional to Simplified");}],
    ["Unblock All My Blocked Areas", function(){ribbon("Review","Block Authors > Unblock All");}],
    ["Undo", function(){shortcut("Ctrl+Z");}],
    ["Ungroup Objects", function(){ribbon("Shape Format","Group > Ungroup");}],
    ["Unlink Subdocument", function(){ribbon("Outlining","Master Document > Unlink");}],
    ["Update All Fields", function(){shortcut("Ctrl+A then F9");}],
    ["Update Bibliography", function(){ribbon("References","Update Citations and Bibliography");}],
    ["Update Field", function(){shortcut("F9 (select field first)");}],
    ["Update Index", function(){ribbon("References","Update Index");}],
    ["Update Style to Match Selection", function(){ribbon("Home","right-click style > Update to Match Selection");}],
    ["Update Table of Contents", function(){ribbon("References","Update Table");}],
    ["Update Table of Figures", function(){ribbon("References","Update Table");}],
    ["View Footnotes", function(){ribbon("References","Show Notes");}],
    ["View Gridlines (Table)", function(){ribbon("Table Layout","View Gridlines");}],
    ["View Macros", function(){shortcut("Alt+F8");}],
    ["View Ruler", function(){ribbon("View","Ruler");}],
    ["View Side by Side", function(){ribbon("View","View Side by Side");}],
    ["Visual Basic Editor", function(){shortcut("Alt+F11");}],
    ["Watermark: Confidential", function(){ribbon("Design","Watermark > Confidential");}],
    ["Watermark: Custom", function(){ribbon("Design","Watermark > Custom Watermark");}],
    ["Watermark: Do Not Copy", function(){ribbon("Design","Watermark > Do Not Copy");}],
    ["Watermark: Draft", function(){ribbon("Design","Watermark > Draft");}],
    ["Watermark: Remove", function(){ribbon("Design","Watermark > Remove Watermark");}],
    ["Watermark: Urgent", function(){ribbon("Design","Watermark > Urgent");}],
    ["Web Layout View", function(){ribbon("View","Web Layout");}],
    ["Whole Page Zoom", function(){ribbon("View","One Page (Zoom group)");}],
    ["Widow/Orphan Control", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Widow/Orphan control");}],
    ["Word Count", function(){ribbon("Review","Word Count");}],
    ["WordArt: Insert", function(){ribbon("Insert","WordArt");}],
    ["Wrap Text", function(){ribbon("Shape Format / Picture Format","Wrap Text");}],
    ["Writing Pen", function(){ribbon("Draw","Pens > Writing Pen");}],
    ["XML Mapping Pane", function(){ribbon("Developer","XML Mapping Pane");}],
    ["XML Schema", function(){ribbon("Developer","XML Schema");}],
    ["Zoom Dialog", function(){ribbon("View","Zoom launcher");}],
    ["Zoom In", function(){shortcut("Ctrl+Mouse Scroll Up or Ctrl++");}],
    ["Zoom Out", function(){shortcut("Ctrl+Mouse Scroll Down or Ctrl+-");}],
    ["Zoom: 100%", function(){ribbon("View","Zoom > 100%");}],
    ["Zoom: 150%", function(){ribbon("View","Zoom > 150%");}],
    ["Zoom: 200%", function(){ribbon("View","Zoom > 200%");}],
    ["Zoom: 50%", function(){ribbon("View","Zoom > 50%");}],
    ["Zoom: 75%", function(){ribbon("View","Zoom > 75%");}],
    ["Zoom: One Page", function(){ribbon("View","Zoom > One Page");}],
    ["Zoom: Page Width", function(){ribbon("View","Zoom > Page Width");}],
    ["Zoom: Two Pages", function(){ribbon("View","Zoom > Two Pages");}],
  ];

  // ── Populate dropdown & wire events ───────────────────────────────
  function setup() {
    var select = document.getElementById("command-list");

    // Build options
    select.innerHTML = "";
    ALL_COMMANDS.forEach(function (entry, idx) {
      var opt = document.createElement("option");
      opt.value = idx;
      opt.textContent = entry[0];
      select.appendChild(opt);
    });

    // Execute immediately when a command is clicked
    select.addEventListener("change", function () {
      var idx = select.value;
      if (idx === "" || idx === undefined) return;
      ALL_COMMANDS[parseInt(idx, 10)][1]();
    });

    // Allow re-clicking the same command
    select.addEventListener("click", function () {
      var idx = select.value;
      if (idx === "" || idx === undefined) return;
      ALL_COMMANDS[parseInt(idx, 10)][1]();
    });
  }

  Office.onReady(function () { setup(); });
})();
