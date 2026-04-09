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
    ["All Caps", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("allCaps");return c.sync().then(function(){s.font.allCaps=!s.font.allCaps;return c.sync();});}).then(function(){showToast("All Caps toggled.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Arrange All Windows", function(){ribbon("View","Arrange All");}],
    ["Attach Template", function(){ribbon("Developer","Document Template");}],
    ["AutoCorrect Options", function(){ribbon("File","Options > Proofing > AutoCorrect Options");}],
    ["AutoFit Contents", function(){ribbon("Table Layout","AutoFit > AutoFit Contents");}],
    ["AutoFit Window", function(){ribbon("Table Layout","AutoFit > AutoFit Window");}],
    ["AutoFormat", function(){ribbon("File","Options > Proofing > AutoFormat");}],
    ["AutoSave", function(){ribbon("File","AutoSave toggle (top-left)");}],
    ["AutoText", function(){ribbon("Insert","Quick Parts > AutoText");}],
    ["Blank Page", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().insertBreak("Page","After");return c.sync();}).then(function(){showToast("Blank page inserted.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Only available in Word.");}
    }],
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
    ["Change Case: UPPERCASE", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){s.insertText(s.text.toUpperCase(),"Replace");return c.sync();});}).then(function(){showToast("Changed to UPPERCASE.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Change Case: lowercase", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){s.insertText(s.text.toLowerCase(),"Replace");return c.sync();});}).then(function(){showToast("Changed to lowercase.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Change Case: Capitalize Each Word", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){s.insertText(s.text.replace(/\b\w/g,function(l){return l.toUpperCase();}),"Replace");return c.sync();});}).then(function(){showToast("Capitalized each word.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Change Case: Sentence case", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){var t=s.text.toLowerCase();t=t.charAt(0).toUpperCase()+t.slice(1);s.insertText(t,"Replace");return c.sync();});}).then(function(){showToast("Sentence case applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Change Case: tOGGLE cASE", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){var t=s.text.split("").map(function(ch){return ch===ch.toUpperCase()?ch.toLowerCase():ch.toUpperCase();}).join("");s.insertText(t,"Replace");return c.sync();});}).then(function(){showToast("Case toggled.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Change Chart Type", function(){ribbon("Chart Design","Change Chart Type");}],
    ["Change Colors (Chart)", function(){ribbon("Chart Design","Change Colors");}],
    ["Change Picture", function(){ribbon("Picture Format","Change Picture");}],
    ["Change Shape", function(){ribbon("Shape Format","Edit Shape > Change Shape");}],
    ["Character Spacing: Expanded", function(){ribbon("Home","Font dialog > Advanced > Spacing > Expanded");}],
    ["Character Spacing: Condensed", function(){ribbon("Home","Font dialog > Advanced > Spacing > Condensed");}],
    ["Check Compatibility", function(){ribbon("File","Info > Check for Issues > Compatibility");}],
    ["Clear All Formatting", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.set({bold:false,italic:false,underline:"None",strikethrough:false,superscript:false,subscript:false,color:"#000000",highlightColor:null});return c.sync();}).then(function(){showToast("Formatting cleared.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Space");}
    }],
    ["Clipboard Pane", function(){shortcut("Ctrl+C twice (or Home > Clipboard launcher)");}],
    ["Close", function(){shortcut("Ctrl+W");}],
    ["Close All", function(){ribbon("File","Close All");}],
    ["Close Header and Footer", function(){ribbon("Header & Footer","Close Header and Footer");}],
    ["Collapse All Headings", function(){ribbon("View","Outline > Collapse All");}],
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
    ["Columns: Two", function(){ribbon("Layout","Columns > Two");}],
    ["Columns: Three", function(){ribbon("Layout","Columns > Three");}],
    ["Combine Documents", function(){ribbon("Review","Compare > Combine");}],
    ["Compare Documents", function(){ribbon("Review","Compare > Compare");}],
    ["Compress Pictures", function(){ribbon("Picture Format","Compress Pictures");}],
    ["Content Control: Plain Text", function(){ribbon("Developer","Plain Text Content Control");}],
    ["Content Control: Rich Text", function(){ribbon("Developer","Rich Text Content Control");}],
    ["Content Control: Picture", function(){ribbon("Developer","Picture Content Control");}],
    ["Content Control: Combo Box", function(){ribbon("Developer","Combo Box Content Control");}],
    ["Content Control: Drop-Down List", function(){ribbon("Developer","Drop-Down List Content Control");}],
    ["Content Control: Date Picker", function(){ribbon("Developer","Date Picker Content Control");}],
    ["Content Control: Check Box", function(){ribbon("Developer","Check Box Content Control");}],
    ["Content Control: Repeating Section", function(){ribbon("Developer","Repeating Section Content Control");}],
    ["Convert Table to Text", function(){ribbon("Table Layout","Convert to Text");}],
    ["Convert Text to Table", function(){ribbon("Insert","Table > Convert Text to Table");}],
    ["Copy", function(){shortcut("Ctrl+C");}],
    ["Cover Page", function(){ribbon("Insert","Cover Page");}],
    ["Create AutoText", function(){shortcut("Alt+F3");}],
    ["Cross-reference", function(){ribbon("References","Cross-reference");}],
    ["Custom Margins", function(){ribbon("Layout","Margins > Custom Margins");}],
    ["Customize Keyboard", function(){ribbon("File","Options > Customize Ribbon > Customize...");}],
    ["Customize Quick Access Toolbar", function(){ribbon("File","Options > Quick Access Toolbar");}],
    ["Customize Ribbon", function(){ribbon("File","Options > Customize Ribbon");}],
    ["Cut", function(){shortcut("Ctrl+X");}],
    ["Date and Time", function(){ribbon("Insert","Date & Time");}],
    ["Decrease Font Size", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("size");return c.sync().then(function(){s.font.size=Math.max(1,s.font.size-1);return c.sync();});}).then(function(){showToast("Font size decreased.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+<");}
    }],
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
    ["Delete Table", function(){ribbon("Table Layout","Delete > Delete Table");}],
    ["Demote (Outline)", function(){ribbon("Outlining","Demote");}],
    ["Demote to Body Text", function(){ribbon("Outlining","Demote to Body Text");}],
    ["Design Mode", function(){ribbon("Developer","Design Mode");}],
    ["Different First Page (Header/Footer)", function(){ribbon("Header & Footer","Different First Page");}],
    ["Different Odd & Even Pages", function(){ribbon("Header & Footer","Different Odd & Even Pages");}],
    ["Distribute Columns Evenly", function(){ribbon("Table Layout","Distribute Columns");}],
    ["Distribute Rows Evenly", function(){ribbon("Table Layout","Distribute Rows");}],
    ["Document Inspector", function(){ribbon("File","Info > Check for Issues > Inspect Document");}],
    ["Document Protection", function(){ribbon("Review","Restrict Editing");}],
    ["Don't Hyphenate", function(){ribbon("Layout","Hyphenation > None");}],
    ["Double Strikethrough", function(){ribbon("Home","Font dialog > Effects > Double Strikethrough");}],
    ["Double Underline", function(){shortcut("Ctrl+Shift+D");}],
    ["Draft View", function(){ribbon("View","Draft");}],
    ["Draw Table", function(){ribbon("Insert","Table > Draw Table");}],
    ["Draw Text Box", function(){ribbon("Insert","Text Box > Draw Text Box");}],
    ["Drawing Canvas", function(){ribbon("Insert","Shapes > New Drawing Canvas");}],
    ["Drop Cap: Dropped", function(){ribbon("Insert","Drop Cap > Dropped");}],
    ["Drop Cap: In Margin", function(){ribbon("Insert","Drop Cap > In Margin");}],
    ["Drop Cap: None", function(){ribbon("Insert","Drop Cap > None");}],
    ["Edit Header", function(){ribbon("Insert","Header > Edit Header");}],
    ["Edit Footer", function(){ribbon("Insert","Footer > Edit Footer");}],
    ["Editing Restrictions", function(){ribbon("Review","Restrict Editing");}],
    ["Editor", function(){ribbon("Home","Editor");}],
    ["Effects (Theme)", function(){ribbon("Design","Effects");}],
    ["Email as Attachment", function(){ribbon("File","Share > Email");}],
    ["Embed Fonts", function(){ribbon("File","Options > Save > Embed Fonts");}],
    ["Enclose Characters", function(){ribbon("Home","Enclose Characters (Asian Layout)");}],
    ["Encrypt with Password", function(){ribbon("File","Info > Protect Document > Encrypt with Password");}],
    ["Endnote: Insert", function(){shortcut("Ctrl+Alt+D");}],
    ["Envelopes", function(){ribbon("Mailings","Envelopes");}],
    ["Equation", function(){shortcut("Alt+=");}],
    ["Eraser (Table)", function(){ribbon("Table Layout","Eraser");}],
    ["Even Page Section Break", function(){wordInsertBreak("EvenPage");}],
    ["Expand All Headings", function(){ribbon("View","Outline > Expand All");}],
    ["Export to PDF/XPS", function(){ribbon("File","Export > Create PDF/XPS");}],
    ["Field", function(){ribbon("Insert","Quick Parts > Field");}],
    ["Field Codes: Toggle", function(){shortcut("Alt+F9");}],
    ["File: New", function(){shortcut("Ctrl+N");}],
    ["File: Open", function(){shortcut("Ctrl+O");}],
    ["File: Print", function(){shortcut("Ctrl+P");}],
    ["File: Save", function(){
      Office.context.document.save(Office.AsyncResultStatus||{},function(r){
        if(r.status===Office.AsyncResultStatus.Failed){showToast("Save failed: "+r.error.message);}else{showToast("Saved!");}
      });
    }],
    ["File: Save As", function(){shortcut("F12 or Ctrl+Shift+S");}],
    ["File: Close", function(){shortcut("Ctrl+W");}],
    ["File: Info", function(){ribbon("File","Info");}],
    ["File: Options", function(){ribbon("File","Options");}],
    ["Find", function(){shortcut("Ctrl+F");}],
    ["Find and Replace", function(){shortcut("Ctrl+H");}],
    ["Find Next", function(){shortcut("Ctrl+G or F5");}],
    ["First Line Indent", function(){setIndent("first",36);}],
    ["Flip Horizontal", function(){ribbon("Shape Format","Rotate > Flip Horizontal");}],
    ["Flip Vertical", function(){ribbon("Shape Format","Rotate > Flip Vertical");}],
    ["Focus Mode", function(){ribbon("View","Focus");}],
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
    ["Font Dialog", function(){shortcut("Ctrl+D");}],
    ["Font Size: 8", function(){setFontSize(8);}],
    ["Font Size: 9", function(){setFontSize(9);}],
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
    ["Footer", function(){ribbon("Insert","Footer");}],
    ["Footnote: Insert", function(){shortcut("Ctrl+Alt+F");}],
    ["Format Painter", function(){shortcut("Ctrl+Shift+C to copy, Ctrl+Shift+V to paste format");}],
    ["Formatting Marks (Show/Hide)", function(){shortcut("Ctrl+Shift+8 or Ctrl+*");}],
    ["Full Screen Reading", function(){ribbon("View","Read Mode");}],
    ["Go Back", function(){shortcut("Alt+Left Arrow");}],
    ["Go Forward", function(){shortcut("Alt+Right Arrow");}],
    ["Go To", function(){shortcut("Ctrl+G or F5");}],
    ["Go to Bookmark", function(){shortcut("Ctrl+Shift+F5");}],
    ["Go to Header", function(){ribbon("Insert","Header > Edit Header");}],
    ["Go to Footer", function(){ribbon("Insert","Footer > Edit Footer");}],
    ["Go to Next Comment", function(){ribbon("Review","Next Comment");}],
    ["Go to Previous Comment", function(){ribbon("Review","Previous Comment");}],
    ["Go to Next Section", function(){ribbon("Navigate: Ctrl+G > Section");}],
    ["Greeting Line (Mail Merge)", function(){ribbon("Mailings","Greeting Line");}],
    ["Gridlines (View)", function(){ribbon("View","Gridlines");}],
    ["Grow Font", function(){shortcut("Ctrl+Shift+>");}],
    ["Group Objects", function(){ribbon("Shape Format","Group > Group");}],
    ["Hanging Indent", function(){setIndent("first",-36);}],
    ["Header", function(){ribbon("Insert","Header");}],
    ["Heading 1 Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 1";return c.sync();}).then(function(){showToast("Heading 1 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+1");}
    }],
    ["Heading 2 Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 2";return c.sync();}).then(function(){showToast("Heading 2 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+2");}
    }],
    ["Heading 3 Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 3";return c.sync();}).then(function(){showToast("Heading 3 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+3");}
    }],
    ["Heading 4 Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 4";return c.sync();}).then(function(){showToast("Heading 4 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Heading 4");}
    }],
    ["Heading 5 Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 5";return c.sync();}).then(function(){showToast("Heading 5 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Heading 5");}
    }],
    ["Heading Rows Repeat", function(){ribbon("Table Layout","Repeat Header Rows");}],
    ["Help", function(){shortcut("F1");}],
    ["Hidden Text", function(){ribbon("Home","Font dialog > Effects > Hidden");}],
    ["Highlight: Yellow", function(){setHighlight("#FFFF00");}],
    ["Highlight: Bright Green", function(){setHighlight("#00FF00");}],
    ["Highlight: Cyan", function(){setHighlight("#00FFFF");}],
    ["Highlight: Pink", function(){setHighlight("#FF00FF");}],
    ["Highlight: Red", function(){setHighlight("#FF0000");}],
    ["Highlight: Blue", function(){setHighlight("#0000FF");}],
    ["Highlight: Dark Blue", function(){setHighlight("#000080");}],
    ["Highlight: Teal", function(){setHighlight("#008080");}],
    ["Highlight: Green", function(){setHighlight("#008000");}],
    ["Highlight: Dark Red", function(){setHighlight("#800000");}],
    ["Highlight: Dark Yellow", function(){setHighlight("#808000");}],
    ["Highlight: Gray 50%", function(){setHighlight("#808080");}],
    ["Highlight: Gray 25%", function(){setHighlight("#C0C0C0");}],
    ["Highlight: Remove", function(){setHighlight(null);}],
    ["Horizontal Line", function(){wordInsertHtml('<hr style="border:1px solid #999;width:100%">',"Horizontal line inserted.");}],
    ["Hyperlink: Insert", function(){shortcut("Ctrl+K");}],
    ["Hyphenation: Automatic", function(){ribbon("Layout","Hyphenation > Automatic");}],
    ["Hyphenation: Manual", function(){ribbon("Layout","Hyphenation > Manual");}],
    ["Hyphenation: None", function(){ribbon("Layout","Hyphenation > None");}],
    ["Icons", function(){ribbon("Insert","Icons");}],
    ["Immersive Reader", function(){ribbon("View","Immersive Reader");}],
    ["Increase Font Size", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("size");return c.sync().then(function(){s.font.size=s.font.size+1;return c.sync();});}).then(function(){showToast("Font size increased.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+>");}
    }],
    ["Increase Indent", function(){setIndent("left",36);}],
    ["Increase List Level", function(){shortcut("Tab in list");}],
    ["Index: Insert", function(){ribbon("References","Insert Index");}],
    ["Index: Mark Entry", function(){shortcut("Alt+Shift+X");}],
    ["Index: Update", function(){ribbon("References","Update Index");}],
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
    ["Insert Table", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().insertTable(3,3,"After",[["","",""],["","",""],["","",""]]);return c.sync();}).then(function(){showToast("Table inserted.");}).catch(function(e){showToast("Error: "+e.message);});}
      else if(isExcel()){Excel.run(function(c){var sh=c.workbook.worksheets.getActiveWorksheet();var r=c.workbook.getSelectedRange();r.load("address");return c.sync().then(function(){sh.tables.add(r.address,true).name="QuickTable_"+Date.now();return c.sync();});}).then(function(){showToast("Table created.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Insert Table of Authorities", function(){ribbon("References","Insert Table of Authorities");}],
    ["Insert Table of Contents", function(){ribbon("References","Table of Contents");}],
    ["Insert Table of Figures", function(){ribbon("References","Insert Table of Figures");}],
    ["Insert Text Box", function(){ribbon("Insert","Text Box");}],
    ["Insert Time", function(){ribbon("Insert","Date & Time (with time format)");}],
    ["Insert WordArt", function(){ribbon("Insert","WordArt");}],
    ["Italic", function(){toggleFont("italic");}],
    ["Justify", function(){setAlignment("Justified");}],
    ["Keep Lines Together", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Keep lines together");}],
    ["Keep with Next", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Keep with next");}],
    ["Labels", function(){ribbon("Mailings","Labels");}],
    ["Landscape Orientation", function(){ribbon("Layout","Orientation > Landscape");}],
    ["Language: Set Proofing", function(){ribbon("Review","Language > Set Proofing Language");}],
    ["Language: Translate Document", function(){ribbon("Review","Translate > Translate Document");}],
    ["Language: Translate Selection", function(){ribbon("Review","Translate > Translate Selection");}],
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
    ["Lock Tracking", function(){ribbon("Review","Track Changes > Lock Tracking");}],
    ["Lowercase", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){s.insertText(s.text.toLowerCase(),"Replace");return c.sync();});}).then(function(){showToast("Lowercase applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{showToast("Not available for this app.");}
    }],
    ["Macros: Record", function(){ribbon("View","Macros > Record Macro");}],
    ["Macros: View", function(){shortcut("Alt+F8");}],
    ["Mail Merge: Start", function(){ribbon("Mailings","Start Mail Merge");}],
    ["Mail Merge: Select Recipients", function(){ribbon("Mailings","Select Recipients");}],
    ["Mail Merge: Preview Results", function(){ribbon("Mailings","Preview Results");}],
    ["Mail Merge: Finish & Merge", function(){ribbon("Mailings","Finish & Merge");}],
    ["Manage Add-ins", function(){ribbon("Insert","My Add-ins > Manage My Add-ins");}],
    ["Manage Sources", function(){ribbon("References","Manage Sources");}],
    ["Manage Styles", function(){ribbon("Home","Styles launcher > Manage Styles");}],
    ["Manual Hyphenation", function(){ribbon("Layout","Hyphenation > Manual");}],
    ["Margins: Normal", function(){ribbon("Layout","Margins > Normal");}],
    ["Margins: Narrow", function(){ribbon("Layout","Margins > Narrow");}],
    ["Margins: Moderate", function(){ribbon("Layout","Margins > Moderate");}],
    ["Margins: Wide", function(){ribbon("Layout","Margins > Wide");}],
    ["Margins: Mirrored", function(){ribbon("Layout","Margins > Mirrored");}],
    ["Mark Citation", function(){ribbon("References","Mark Citation");}],
    ["Mark Index Entry", function(){shortcut("Alt+Shift+X");}],
    ["Mark Table of Contents Entry", function(){ribbon("References","Add Text");}],
    ["Merge Cells", function(){ribbon("Table Layout","Merge Cells");}],
    ["Merge Formatting (Paste)", function(){shortcut("Ctrl+Shift+V (then choose)");}],
    ["Modify Style", function(){ribbon("Home","Styles > right-click style > Modify");}],
    ["Move Down (Outline)", function(){ribbon("Outlining","Move Down");}],
    ["Move Up (Outline)", function(){ribbon("Outlining","Move Up");}],
    ["Multilevel List", function(){ribbon("Home","Multilevel List");}],
    ["Navigation Pane", function(){shortcut("Ctrl+F (opens Navigation)");}],
    ["New Blank Document", function(){shortcut("Ctrl+N");}],
    ["New Comment", function(){shortcut("Ctrl+Alt+M");}],
    ["New Folder", function(){ribbon("File","Save As > New Folder");}],
    ["New from Template", function(){ribbon("File","New");}],
    ["New Window", function(){ribbon("View","New Window");}],
    ["Next Change", function(){ribbon("Review","Next Change");}],
    ["Next Comment", function(){ribbon("Review","Next Comment");}],
    ["Next Footnote", function(){ribbon("References","Next Footnote");}],
    ["Next Page Section Break", function(){wordInsertBreak("SectionNext");}],
    ["No Spacing Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="No Spacing";return c.sync();}).then(function(){showToast("No Spacing applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > No Spacing");}
    }],
    ["Normal Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Normal";return c.sync();}).then(function(){showToast("Normal style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+N");}
    }],
    ["Numbering", function(){ribbon("Home","Numbering");}],
    ["Object: Insert", function(){ribbon("Insert","Object");}],
    ["Odd Page Section Break", function(){wordInsertBreak("OddPage");}],
    ["Open", function(){shortcut("Ctrl+O");}],
    ["Open Hyperlink", function(){shortcut("Ctrl+Click on hyperlink");}],
    ["Open in Browser", function(){ribbon("File","Info > Open in Browser");}],
    ["Orientation: Portrait", function(){ribbon("Layout","Orientation > Portrait");}],
    ["Orientation: Landscape", function(){ribbon("Layout","Orientation > Landscape");}],
    ["Outline View", function(){ribbon("View","Outline");}],
    ["Page Background: Color", function(){ribbon("Design","Page Color");}],
    ["Page Background: Watermark", function(){ribbon("Design","Watermark");}],
    ["Page Background: Page Borders", function(){ribbon("Design","Page Borders");}],
    ["Page Break", function(){wordInsertBreak("Page");}],
    ["Page Break Before", function(){ribbon("Home","Paragraph dialog > Line and Page Breaks > Page break before");}],
    ["Page Color", function(){ribbon("Design","Page Color");}],
    ["Page Down", function(){shortcut("Page Down key");}],
    ["Page Layout View", function(){ribbon("View","Print Layout");}],
    ["Page Number: Top of Page", function(){ribbon("Insert","Page Number > Top of Page");}],
    ["Page Number: Bottom of Page", function(){ribbon("Insert","Page Number > Bottom of Page");}],
    ["Page Number: Page Margins", function(){ribbon("Insert","Page Number > Page Margins");}],
    ["Page Number: Current Position", function(){ribbon("Insert","Page Number > Current Position");}],
    ["Page Number: Format", function(){ribbon("Insert","Page Number > Format Page Numbers");}],
    ["Page Number: Remove", function(){ribbon("Insert","Page Number > Remove Page Numbers");}],
    ["Page Setup Dialog", function(){ribbon("Layout","Page Setup launcher (bottom-right arrow)");}],
    ["Page Up", function(){shortcut("Page Up key");}],
    ["Paragraph Dialog", function(){ribbon("Home","Paragraph launcher (bottom-right arrow)");}],
    ["Paragraph Marks (Show/Hide)", function(){shortcut("Ctrl+Shift+8");}],
    ["Paragraph Shading", function(){ribbon("Home","Shading (paint bucket)");}],
    ["Paste", function(){shortcut("Ctrl+V");}],
    ["Paste as Hyperlink", function(){ribbon("Home","Paste > Paste Special > Paste as Hyperlink");}],
    ["Paste as Picture", function(){ribbon("Home","Paste > Paste Special > Picture");}],
    ["Paste: Keep Source Formatting", function(){shortcut("Ctrl+V then Ctrl > K");}],
    ["Paste: Keep Text Only", function(){shortcut("Ctrl+V then Ctrl > T");}],
    ["Paste: Merge Formatting", function(){shortcut("Ctrl+V then Ctrl > M");}],
    ["Paste Special", function(){shortcut("Ctrl+Alt+V");}],
    ["Picture: Insert", function(){ribbon("Insert","Pictures");}],
    ["Picture: Border", function(){ribbon("Picture Format","Picture Border");}],
    ["Picture: Effects", function(){ribbon("Picture Format","Picture Effects");}],
    ["Picture: Layout", function(){ribbon("Picture Format","Picture Layout");}],
    ["Picture: Compress", function(){ribbon("Picture Format","Compress Pictures");}],
    ["Picture: Reset", function(){ribbon("Picture Format","Reset Picture");}],
    ["Picture: Crop", function(){ribbon("Picture Format","Crop");}],
    ["Portrait Orientation", function(){ribbon("Layout","Orientation > Portrait");}],
    ["Position Object", function(){ribbon("Shape Format","Position");}],
    ["Present Online", function(){ribbon("File","Share > Present Online");}],
    ["Previous Change", function(){ribbon("Review","Previous Change");}],
    ["Previous Comment", function(){ribbon("Review","Previous Comment");}],
    ["Print", function(){shortcut("Ctrl+P");}],
    ["Print Layout View", function(){ribbon("View","Print Layout");}],
    ["Print Preview", function(){shortcut("Ctrl+F2");}],
    ["Promote (Outline)", function(){ribbon("Outlining","Promote");}],
    ["Protect Document", function(){ribbon("File","Info > Protect Document");}],
    ["Quick Parts", function(){ribbon("Insert","Quick Parts");}],
    ["Quick Print", function(){shortcut("Ctrl+P then Enter");}],
    ["Quick Styles", function(){ribbon("Home","Styles gallery");}],
    ["Quick Tables", function(){ribbon("Insert","Table > Quick Tables");}],
    ["Read Mode", function(){ribbon("View","Read Mode");}],
    ["Redo", function(){shortcut("Ctrl+Y");}],
    ["References: Bibliography", function(){ribbon("References","Bibliography");}],
    ["References: Manage Sources", function(){ribbon("References","Manage Sources");}],
    ["References: Style (APA, MLA, etc.)", function(){ribbon("References","Style dropdown");}],
    ["Reject All Changes in Document", function(){ribbon("Review","Reject > Reject All Changes");}],
    ["Reject and Move to Next", function(){ribbon("Review","Reject");}],
    ["Remove All Formatting", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.set({bold:false,italic:false,underline:"None",strikethrough:false,superscript:false,subscript:false,color:"#000000",highlightColor:null});return c.sync();}).then(function(){showToast("Formatting removed.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Space");}
    }],
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
    ["Reviewing Pane: Horizontal", function(){ribbon("Review","Reviewing Pane > Horizontal");}],
    ["Reviewing Pane: Vertical", function(){ribbon("Review","Reviewing Pane > Vertical");}],
    ["Right Indent", function(){setIndent("right",36);}],
    ["Right Tab Stop", function(){ribbon("Home","Paragraph dialog > Tabs");}],
    ["Rotate Left 90", function(){ribbon("Shape Format","Rotate > Rotate Left 90");}],
    ["Rotate Right 90", function(){ribbon("Shape Format","Rotate > Rotate Right 90");}],
    ["Ruler (Show/Hide)", function(){ribbon("View","Ruler");}],
    ["Save", function(){
      Office.context.document.save(Office.AsyncResultStatus||{},function(r){
        if(r.status===Office.AsyncResultStatus.Failed){showToast("Save failed: "+r.error.message);}else{showToast("Saved!");}
      });
    }],
    ["Save All", function(){shortcut("Not available via API. Save each document individually.");}],
    ["Save As", function(){shortcut("F12 or Ctrl+Shift+S");}],
    ["Save As PDF", function(){ribbon("File","Export > Create PDF/XPS");}],
    ["Save As Template", function(){ribbon("File","Save As > Word Template (*.dotx)");}],
    ["Screenshot", function(){ribbon("Insert","Screenshot");}],
    ["Screen Clipping", function(){ribbon("Insert","Screenshot > Screen Clipping");}],
    ["Search Document", function(){shortcut("Ctrl+F");}],
    ["Section Break: Continuous", function(){wordInsertBreak("SectionContinuous");}],
    ["Section Break: Even Page", function(){wordInsertBreak("EvenPage");}],
    ["Section Break: Next Page", function(){wordInsertBreak("SectionNext");}],
    ["Section Break: Odd Page", function(){wordInsertBreak("OddPage");}],
    ["Select All", function(){
      if(isWord()){Word.run(function(c){c.document.body.getRange().select();return c.sync();}).then(function(){showToast("All selected.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+A");}
    }],
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
    ["Shading", function(){ribbon("Home","Shading (paint bucket icon)");}],
    ["Shape Effects", function(){ribbon("Shape Format","Shape Effects");}],
    ["Shape Fill", function(){ribbon("Shape Format","Shape Fill");}],
    ["Shape Outline", function(){ribbon("Shape Format","Shape Outline");}],
    ["Shapes: Insert", function(){ribbon("Insert","Shapes");}],
    ["Share", function(){ribbon("File","Share");}],
    ["Show All Formatting Marks", function(){shortcut("Ctrl+Shift+8 or Ctrl+*");}],
    ["Show Comments", function(){ribbon("Review","Show Comments");}],
    ["Show Markup", function(){ribbon("Review","Show Markup");}],
    ["Show/Hide Paragraph Marks", function(){shortcut("Ctrl+Shift+8");}],
    ["Shrink Font", function(){shortcut("Ctrl+Shift+<");}],
    ["Shrink One Page", function(){ribbon("Print Preview","Shrink One Page");}],
    ["Simple Markup", function(){ribbon("Review","Display for Review > Simple Markup");}],
    ["Size: Letter (8.5x11)", function(){ribbon("Layout","Size > Letter");}],
    ["Size: Legal (8.5x14)", function(){ribbon("Layout","Size > Legal");}],
    ["Size: A4", function(){ribbon("Layout","Size > A4");}],
    ["Size: A5", function(){ribbon("Layout","Size > A5");}],
    ["Size: B5", function(){ribbon("Layout","Size > B5");}],
    ["Size: Executive", function(){ribbon("Layout","Size > Executive");}],
    ["Small Caps", function(){shortcut("Ctrl+Shift+K");}],
    ["SmartArt", function(){ribbon("Insert","SmartArt");}],
    ["Snap to Grid", function(){ribbon("Shape Format","Align > Grid Settings > Snap to Grid");}],
    ["Sort", function(){ribbon("Home","Sort (A-Z icon)");}],
    ["Sort Ascending", function(){ribbon("Home","Sort > Ascending");}],
    ["Sort Descending", function(){ribbon("Home","Sort > Descending");}],
    ["Space After: 0 pt", function(){setSpaceAfter(0);}],
    ["Space After: 6 pt", function(){setSpaceAfter(6);}],
    ["Space After: 8 pt", function(){setSpaceAfter(8);}],
    ["Space After: 10 pt", function(){setSpaceAfter(10);}],
    ["Space After: 12 pt", function(){setSpaceAfter(12);}],
    ["Space After: 24 pt", function(){setSpaceAfter(24);}],
    ["Space Before: 0 pt", function(){setSpaceBefore(0);}],
    ["Space Before: 6 pt", function(){setSpaceBefore(6);}],
    ["Space Before: 8 pt", function(){setSpaceBefore(8);}],
    ["Space Before: 10 pt", function(){setSpaceBefore(10);}],
    ["Space Before: 12 pt", function(){setSpaceBefore(12);}],
    ["Space Before: 24 pt", function(){setSpaceBefore(24);}],
    ["Spelling and Grammar", function(){shortcut("F7");}],
    ["Split Cells", function(){ribbon("Table Layout","Split Cells");}],
    ["Split Table", function(){ribbon("Table Layout","Split Table");}],
    ["Split Window", function(){ribbon("View","Split");}],
    ["Start Mail Merge", function(){ribbon("Mailings","Start Mail Merge");}],
    ["Start Tracking Changes", function(){shortcut("Ctrl+Shift+E");}],
    ["Strikethrough", function(){toggleFont("strikethrough");}],
    ["Style: Heading 1", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 1";return c.sync();}).then(function(){showToast("Heading 1 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+1");}
    }],
    ["Style: Heading 2", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 2";return c.sync();}).then(function(){showToast("Heading 2 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+2");}
    }],
    ["Style: Heading 3", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Heading 3";return c.sync();}).then(function(){showToast("Heading 3 applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Alt+3");}
    }],
    ["Style: Normal", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Normal";return c.sync();}).then(function(){showToast("Normal style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+N");}
    }],
    ["Style: No Spacing", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="No Spacing";return c.sync();}).then(function(){showToast("No Spacing applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > No Spacing");}
    }],
    ["Style: Title", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Title";return c.sync();}).then(function(){showToast("Title style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Title");}
    }],
    ["Style: Subtitle", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Subtitle";return c.sync();}).then(function(){showToast("Subtitle style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Subtitle");}
    }],
    ["Style: Emphasis", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Emphasis";return c.sync();}).then(function(){showToast("Emphasis applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Emphasis");}
    }],
    ["Style: Strong", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Strong";return c.sync();}).then(function(){showToast("Strong applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Strong");}
    }],
    ["Style: Quote", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Quote";return c.sync();}).then(function(){showToast("Quote style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Quote");}
    }],
    ["Style: Intense Quote", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Intense Quote";return c.sync();}).then(function(){showToast("Intense Quote applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Intense Quote");}
    }],
    ["Style: List Paragraph", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="List Paragraph";return c.sync();}).then(function(){showToast("List Paragraph applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > List Paragraph");}
    }],
    ["Style: Subtle Emphasis", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Subtle Emphasis";return c.sync();}).then(function(){showToast("Subtle Emphasis applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Subtle Emphasis");}
    }],
    ["Style: Intense Emphasis", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Intense Emphasis";return c.sync();}).then(function(){showToast("Intense Emphasis applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Intense Emphasis");}
    }],
    ["Style: Subtle Reference", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Subtle Reference";return c.sync();}).then(function(){showToast("Subtle Reference applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Subtle Reference");}
    }],
    ["Style: Intense Reference", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Intense Reference";return c.sync();}).then(function(){showToast("Intense Reference applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Intense Reference");}
    }],
    ["Style: Book Title", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Book Title";return c.sync();}).then(function(){showToast("Book Title applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Book Title");}
    }],
    ["Subscript", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("subscript");return c.sync().then(function(){s.font.subscript=!s.font.subscript;return c.sync();});}).then(function(){showToast("Subscript toggled.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+=");}
    }],
    ["Superscript", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("superscript");return c.sync().then(function(){s.font.superscript=!s.font.superscript;return c.sync();});}).then(function(){showToast("Superscript toggled.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift++");}
    }],
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
    ["Table of Contents: Update", function(){ribbon("References","Update Table");}],
    ["Table of Contents: Remove", function(){ribbon("References","Table of Contents > Remove");}],
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
    ["Text Wrapping: In Line with Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > In Line with Text");}],
    ["Text Wrapping: Square", function(){ribbon("Shape Format/Picture Format","Wrap Text > Square");}],
    ["Text Wrapping: Tight", function(){ribbon("Shape Format/Picture Format","Wrap Text > Tight");}],
    ["Text Wrapping: Through", function(){ribbon("Shape Format/Picture Format","Wrap Text > Through");}],
    ["Text Wrapping: Top and Bottom", function(){ribbon("Shape Format/Picture Format","Wrap Text > Top and Bottom");}],
    ["Text Wrapping: Behind Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > Behind Text");}],
    ["Text Wrapping: In Front of Text", function(){ribbon("Shape Format/Picture Format","Wrap Text > In Front of Text");}],
    ["Theme Colors", function(){ribbon("Design","Colors");}],
    ["Theme Effects", function(){ribbon("Design","Effects");}],
    ["Theme Fonts", function(){ribbon("Design","Fonts");}],
    ["Themes", function(){ribbon("Design","Themes");}],
    ["Thesaurus", function(){shortcut("Shift+F7");}],
    ["Title Style", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().style="Title";return c.sync();}).then(function(){showToast("Title style applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Styles > Title");}
    }],
    ["Toggle Field Codes", function(){shortcut("Alt+F9");}],
    ["Toggle Full Screen", function(){shortcut("Alt+V, U (or View > Full Screen)");}],
    ["Track Changes", function(){shortcut("Ctrl+Shift+E");}],
    ["Track Changes Options", function(){ribbon("Review","Track Changes > Change Tracking Options");}],
    ["Translate Document", function(){ribbon("Review","Translate > Translate Document");}],
    ["Translate Selection", function(){ribbon("Review","Translate > Translate Selection");}],
    ["Underline", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.font.load("underline");return c.sync().then(function(){s.font.underline=s.font.underline==="None"?"Single":"None";return c.sync();});}).then(function(){showToast("Underline toggled.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{toggleFont("underline");}
    }],
    ["Underline: Double", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="Double";return c.sync();}).then(function(){showToast("Double underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+D");}
    }],
    ["Underline: Dotted", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="Dotted";return c.sync();}).then(function(){showToast("Dotted underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Font dialog > Underline style");}
    }],
    ["Underline: Dashed", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="DashLine";return c.sync();}).then(function(){showToast("Dashed underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Font dialog > Underline style");}
    }],
    ["Underline: Wavy", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="Wave";return c.sync();}).then(function(){showToast("Wavy underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Font dialog > Underline style");}
    }],
    ["Underline: Thick", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="Thick";return c.sync();}).then(function(){showToast("Thick underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{ribbon("Home","Font dialog > Underline style");}
    }],
    ["Underline: Word Only", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="Word";return c.sync();}).then(function(){showToast("Word-only underline applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+W");}
    }],
    ["Underline: None", function(){
      if(isWord()){Word.run(function(c){c.document.getSelection().font.underline="None";return c.sync();}).then(function(){showToast("Underline removed.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+U to toggle");}
    }],
    ["Undo", function(){shortcut("Ctrl+Z");}],
    ["Ungroup Objects", function(){ribbon("Shape Format","Group > Ungroup");}],
    ["Update All Fields", function(){shortcut("Ctrl+A then F9");}],
    ["Update Bibliography", function(){ribbon("References","Update Citations and Bibliography");}],
    ["Update Field", function(){shortcut("F9 (select field first)");}],
    ["Update Index", function(){ribbon("References","Update Index");}],
    ["Update Style to Match Selection", function(){ribbon("Home","right-click style > Update to Match Selection");}],
    ["Update Table of Contents", function(){ribbon("References","Update Table");}],
    ["Update Table of Figures", function(){ribbon("References","Update Table");}],
    ["UPPERCASE", function(){
      if(isWord()){Word.run(function(c){var s=c.document.getSelection();s.load("text");return c.sync().then(function(){s.insertText(s.text.toUpperCase(),"Replace");return c.sync();});}).then(function(){showToast("UPPERCASE applied.");}).catch(function(e){showToast("Error: "+e.message);});}
      else{shortcut("Ctrl+Shift+A (in Word)");}
    }],
    ["View Gridlines (Table)", function(){ribbon("Table Layout","View Gridlines");}],
    ["View Footnotes", function(){ribbon("References","Show Notes");}],
    ["View Macros", function(){shortcut("Alt+F8");}],
    ["View Ruler", function(){ribbon("View","Ruler");}],
    ["View Side by Side", function(){ribbon("View","View Side by Side");}],
    ["Visual Basic Editor", function(){shortcut("Alt+F11");}],
    ["Watermark: Custom", function(){ribbon("Design","Watermark > Custom Watermark");}],
    ["Watermark: Confidential", function(){ribbon("Design","Watermark > Confidential");}],
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
    ["XML Mapping Pane", function(){ribbon("Developer","XML Mapping Pane");}],
    ["XML Schema", function(){ribbon("Developer","XML Schema");}],
    ["Zoom: 50%", function(){ribbon("View","Zoom > 50%");}],
    ["Zoom: 75%", function(){ribbon("View","Zoom > 75%");}],
    ["Zoom: 100%", function(){ribbon("View","Zoom > 100%");}],
    ["Zoom: 150%", function(){ribbon("View","Zoom > 150%");}],
    ["Zoom: 200%", function(){ribbon("View","Zoom > 200%");}],
    ["Zoom: One Page", function(){ribbon("View","Zoom > One Page");}],
    ["Zoom: Two Pages", function(){ribbon("View","Zoom > Two Pages");}],
    ["Zoom: Page Width", function(){ribbon("View","Zoom > Page Width");}],
    ["Zoom Dialog", function(){ribbon("View","Zoom launcher");}],
    ["Zoom In", function(){shortcut("Ctrl+Mouse Scroll Up or Ctrl++");}],
    ["Zoom Out", function(){shortcut("Ctrl+Mouse Scroll Down or Ctrl+-");}],
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
