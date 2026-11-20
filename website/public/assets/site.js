/* vba-edit docs: theme, "your setup" switcher, terminal player, tabs, table of contents. No dependencies. */
(function () {
  "use strict";

  var store = {
    get: function (key) { try { return localStorage.getItem(key); } catch (e) { return null; } },
    set: function (key, value) { try { localStorage.setItem(key, value); } catch (e) { /* private mode */ } },
  };
  var reducedMotion = window.matchMedia && window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  function $(selector, root) { return (root || document).querySelector(selector); }
  function $$(selector, root) { return Array.prototype.slice.call((root || document).querySelectorAll(selector)); }
  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text != null) node.textContent = text;
    return node;
  }

  /* ── Theme ───────────────────────────────────────────────────────────── */
  function initTheme() {
    var button = $("[data-theme-toggle]");
    if (!button) return;
    button.addEventListener("click", function () {
      var root = document.documentElement;
      var dark = root.dataset.theme
        ? root.dataset.theme === "dark"
        : window.matchMedia("(prefers-color-scheme: dark)").matches;
      root.dataset.theme = dark ? "light" : "dark";
      store.set("vba-edit-theme", root.dataset.theme);
    });
  }

  /* ── Your setup: which Office app, and how the tool is started ───────── */
  var TOOLS = {
    excel: { tool: "excel-vba", doc: "Budget.xlsm", noun: "workbook", app: "Excel" },
    word: { tool: "word-vba", doc: "Letter.docm", noun: "document", app: "Word" },
    powerpoint: { tool: "powerpoint-vba", doc: "Deck.pptm", noun: "presentation", app: "PowerPoint" },
    access: { tool: "access-vba", doc: "Orders.accdb", noun: "database", app: "Access" },
  };
  /* What is typed in front of (and behind) the tool name. A downloaded .exe is not on the PATH, and
     PowerShell only starts a program from the current folder when it is asked to with ".\". */
  var RUNNERS = {
    installed: { before: "", after: "" },
    uvx: { before: "uvx ", after: "" },
    exe: { before: ".\\", after: ".exe" },
  };
  /* ?app=word&run=exe in the address selects a setup (and makes it linkable); otherwise the last choice. */
  var query = {};
  try {
    new URLSearchParams(window.location.search).forEach(function (value, key) { query[key] = value; });
  } catch (e) { /* very old browser: no query support */ }
  var setup = {
    tool: query.app || store.get("vba-edit-tool") || "excel",
    run: query.run || store.get("vba-edit-run") || "installed",
  };
  if (!Object.prototype.hasOwnProperty.call(TOOLS, setup.tool)) setup.tool = "excel";
  if (!Object.prototype.hasOwnProperty.call(RUNNERS, setup.run)) setup.run = "installed";
  var tokens = [];

  /* Samples are written for Excel. Wrap the things that differ per app so they can be swapped: the tool,
     the document, and (marked in the page with data-doc-stem / data-app-noun / data-app-name) the names
     derived from them. A config file shown next to a command is swapped with it; only blocks that are
     about one particular application are marked data-fixed. */
  function tokenize() {
    var pattern = /excel-vba|Budget\.xlsm/g;
    $$("pre, code, .card-cmd, .synopsis").forEach(function (block) {
      if (block.closest(".term, [data-fixed], .editor")) return;
      var walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
      var nodes = [];
      while (walker.nextNode()) nodes.push(walker.currentNode);
      nodes.forEach(function (node) {
        if (node.parentNode.classList && node.parentNode.classList.contains("tk")) return;
        var text = node.nodeValue;
        if (!pattern.test(text)) return;
        pattern.lastIndex = 0;
        var fragment = document.createDocumentFragment();
        var last = 0;
        var match;
        while ((match = pattern.exec(text))) {
          fragment.appendChild(document.createTextNode(text.slice(last, match.index)));
          var span = el("span", "tk", match[0]);
          span.dataset.tk = match[0] === "excel-vba" ? "tool" : "doc";
          tokens.push(span);
          fragment.appendChild(span);
          last = match.index + match[0].length;
        }
        fragment.appendChild(document.createTextNode(text.slice(last)));
        node.parentNode.replaceChild(fragment, node);
      });
    });
  }

  function applySetup(flash) {
    var chosen = TOOLS[setup.tool];
    tokens.forEach(function (span) {
      var value;
      if (span.dataset.tk === "doc") value = chosen.doc;
      else value = RUNNERS[setup.run].before + chosen.tool + RUNNERS[setup.run].after;
      if (span.textContent !== value) {
        span.textContent = value;
        if (flash && !reducedMotion) {
          span.classList.add("flash");
          setTimeout(function () { span.classList.remove("flash"); }, 700);
        }
      }
    });
    $$("[data-tool]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.tool === setup.tool)); });
    $$("[data-run]").forEach(function (b) { b.setAttribute("aria-pressed", String(b.dataset.run === setup.run)); });
    $$("[data-app-name]").forEach(function (n) { n.textContent = chosen.app; });
    $$("[data-app-noun]").forEach(function (n) { n.textContent = chosen.noun; });
    $$("[data-doc-stem]").forEach(function (n) { n.textContent = chosen.doc.replace(/\.[^.]+$/, ""); });
    $$("[data-only]").forEach(function (n) { n.hidden = n.dataset.only.split(" ").indexOf(setup.tool) === -1; });
  }

  function initSetup() {
    tokenize();
    $$("[data-tool]").forEach(function (button) {
      button.addEventListener("click", function () {
        setup.tool = button.dataset.tool;
        store.set("vba-edit-tool", setup.tool);
        applySetup(true);
      });
    });
    $$("[data-run]").forEach(function (button) {
      button.addEventListener("click", function () {
        setup.run = button.dataset.run;
        store.set("vba-edit-run", setup.run);
        applySetup(true);
      });
    });
    applySetup(false);
  }

  /* ── Copy buttons ────────────────────────────────────────────────────── */
  function initCopy() {
    $$("pre.cmd, pre[data-copy]").forEach(function (pre) {
      var box = pre.parentNode.classList.contains("codebox") ? pre.parentNode : null;
      if (!box) {
        box = el("div", "codebox");
        pre.parentNode.insertBefore(box, pre);
        box.appendChild(pre);
      }
      if (!navigator.clipboard) return;
      var button = el("button", "copy", "copy");
      button.type = "button";
      button.setAttribute("aria-label", "Copy to clipboard");
      button.addEventListener("click", function () {
        var clone = pre.cloneNode(true);
        $$(".p, .c", clone).forEach(function (n) { n.remove(); });
        var text = clone.textContent.split("\n").map(function (l) { return l.replace(/\s+$/, ""); })
          .filter(function (l, i, all) { return l || (i && all[i - 1]); }).join("\n").trim();
        navigator.clipboard.writeText(text).then(function () {
          button.textContent = "copied";
          button.classList.add("done");
          setTimeout(function () { button.textContent = "copy"; button.classList.remove("done"); }, 1400);
        });
      });
      box.appendChild(button);
    });
  }

  /* ── Tabs ────────────────────────────────────────────────────────────── */
  function initTabs() {
    $$(".tabs").forEach(function (tabs) {
      var buttons = $$(".tablist button", tabs);
      var panels = $$(".tabpanel", tabs);
      function select(index, focus) {
        buttons.forEach(function (b, i) {
          b.setAttribute("aria-selected", String(i === index));
          b.tabIndex = i === index ? 0 : -1;
          panels[i].hidden = i !== index;
        });
        if (focus) buttons[index].focus();
      }
      buttons.forEach(function (button, index) {
        button.setAttribute("role", "tab");
        button.addEventListener("click", function () { select(index); });
        button.addEventListener("keydown", function (event) {
          if (event.key === "ArrowRight") select((index + 1) % buttons.length, true);
          if (event.key === "ArrowLeft") select((index + buttons.length - 1) % buttons.length, true);
        });
      });
      select(0);
    });
  }

  /* ── Terminal player ─────────────────────────────────────────────────── */
  var PROMPT = "PS C:\\Projects\\budget> ";

  /* Output arrives as HTML with flat <span> runs; make every line stand on its own. */
  function splitLines(html) {
    var open = "";
    return html.split("\n").map(function (line) {
      var full = open + line;
      var lastOpen = full.lastIndexOf("<span");
      var lastClose = full.lastIndexOf("</span>");
      if (lastOpen > lastClose) {
        open = full.slice(lastOpen, full.indexOf(">", lastOpen) + 1);
        return full + "</span>";
      }
      open = "";
      return full;
    });
  }

  function Player(root, steps) {
    this.root = root;
    this.steps = steps;
    this.run = 0;
    this.body = $(".term-body", root);
    this.fallback = this.body.innerHTML;
    var self = this;
    var controls = el("div", "term-ctl");
    var replay = el("button", null, "replay");
    var skip = el("button", null, "skip");
    replay.type = skip.type = "button";
    replay.addEventListener("click", function () { self.play(); });
    skip.addEventListener("click", function () { self.finish(); });
    controls.appendChild(skip);
    controls.appendChild(replay);
    $(".term-bar", root).appendChild(controls);
    this.skipButton = skip;
  }

  Player.prototype.line = function (className) {
    var line = el("div", className || "tl");
    this.body.appendChild(line);
    return line;
  };

  Player.prototype.scroll = function () { this.body.scrollTop = this.body.scrollHeight; };

  /* Build every step as a list of small actions, each with the pause that follows it. */
  Player.prototype.actions = function () {
    var self = this;
    var list = [];
    var promptLine = null;
    function add(delay, action) { list.push({ delay: delay, action: action }); }

    this.steps.forEach(function (step) {
      if (step.cmd != null) {
        var typed;
        add(380, function () {
          var line = self.line();
          line.appendChild(el("span", "tp", PROMPT));
          typed = el("span", "tc");
          line.appendChild(typed);
          line.appendChild(el("span", "tcur"));
        });
        var per = Math.max(9, Math.min(26, 1500 / step.cmd.length));
        step.cmd.split("").forEach(function (character) {
          add(per, function () { typed.textContent += character; });
        });
        add(420, function () { var cursor = $(".tcur", self.body); if (cursor) cursor.remove(); });
      } else if (step.out != null) {
        var lines = splitLines(step.out);
        var each = Math.max(12, Math.min(60, 1300 / lines.length));
        lines.forEach(function (html, index) {
          add(each, function () {
            var line = self.line();
            line.innerHTML = html;
            if (step.prompt && index === lines.length - 1) {
              promptLine = line;
              line.appendChild(el("span", "tcur"));
            }
          });
        });
        if (step.prompt) add(700, function () {});
      } else if (step.typed != null) {
        add(500, function () {
          var cursor = promptLine && $(".tcur", promptLine);
          if (cursor) cursor.remove();
          /* The recorder trims trailing blanks, so put back the space that follows the prompt. */
          if (promptLine) promptLine.appendChild(el("span", "tc", " " + step.typed));
        });
      } else if (step.note != null) {
        add(250, function () {});
        add(1500, function () { self.line("tn").textContent = step.note; });
      } else if (step.key != null) {
        add(600, function () { self.line().appendChild(el("span", "tk-key", "^C")); });
      } else if (step.exit != null) {
        add(500, function () { self.line("tn exit").textContent = "exit code " + step.exit; });
      }
    });
    /* A slice that ends while the command is still running gets a bare cursor instead of a new prompt. */
    add(0, function () {
      var line = self.line();
      if (!self.root.hasAttribute("data-running")) line.appendChild(el("span", "tp", PROMPT));
      line.appendChild(el("span", "tcur"));
    });
    return list;
  };

  Player.prototype.play = function () {
    var self = this;
    var run = ++this.run;
    var list = this.actions();
    var index = 0;
    this.body.textContent = "";
    this.skipButton.hidden = false;
    (function next() {
      if (run !== self.run) return;
      if (index >= list.length) { self.skipButton.hidden = true; return; }
      var item = list[index++];
      item.action();
      self.scroll();
      setTimeout(next, item.delay);
    })();
  };

  Player.prototype.finish = function () {
    this.run++;
    this.body.textContent = "";
    this.actions().forEach(function (item) { item.action(); });
    $$(".tcur", this.body).slice(0, -1).forEach(function (cursor) { cursor.remove(); });
    this.skipButton.hidden = true;
    this.body.scrollTop = 0;
  };

  function initTerminals() {
    var casts = window.VBA_EDIT_CASTS || {};
    var players = [];
    $$(".term[data-cast]").forEach(function (root) {
      var steps = casts[root.dataset.cast];
      if (!steps) return;
      if (root.dataset.steps) {
        var range = root.dataset.steps.split("-").map(Number);
        steps = steps.slice(range[0], range[1] == null ? undefined : range[1]);
      }
      var player = new Player(root, steps);
      players.push(player);
      if (reducedMotion || !("IntersectionObserver" in window)) { player.finish(); return; }
      player.finish(); /* keeps the layout stable until the animation starts */
      player.started = false;
    });
    if (reducedMotion || !("IntersectionObserver" in window)) return;
    var observer = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        var player = players.filter(function (p) { return p.root === entry.target; })[0];
        if (player && !player.started) { player.started = true; player.play(); }
        observer.unobserve(entry.target);
      });
    }, { threshold: 0.45 });
    players.forEach(function (player) { observer.observe(player.root); });
  }

  /* ── Screenshots: click to enlarge ───────────────────────────────────── */
  function initLightbox() {
    var shots = $$(".shot button");
    if (!shots.length || !window.HTMLDialogElement) return;
    var dialog = el("dialog", "lightbox");
    var image = el("img");
    dialog.appendChild(image);
    document.body.appendChild(dialog);
    dialog.addEventListener("click", function () { dialog.close(); });
    shots.forEach(function (button) {
      button.addEventListener("click", function () {
        var source = $("img", button);
        image.src = source.src;
        image.alt = source.alt;
        dialog.showModal();
      });
    });
  }

  /* ── Guide: heading anchors, scroll spy, filter ──────────────────────── */
  function initGuide() {
    var toc = $(".toc");
    if (!toc) return;
    $$(".prose h2[id], .prose h3[id]").forEach(function (heading) {
      var anchor = el("a", "anchor", "#");
      anchor.href = "#" + heading.id;
      anchor.setAttribute("aria-label", "Link to this section");
      heading.appendChild(anchor);
    });

    var links = $$("a[href^='#']", toc);
    var targets = links.map(function (link) { return document.getElementById(link.getAttribute("href").slice(1)); });
    function spy() {
      var line = window.scrollY + 120;
      var current = 0;
      targets.forEach(function (target, index) { if (target && target.offsetTop <= line) current = index; });
      links.forEach(function (link, index) { link.classList.toggle("on", index === current); });
      var active = links[current];
      if (active && toc.scrollHeight > toc.clientHeight) {
        var top = active.offsetTop - toc.clientHeight / 2;
        if (Math.abs(toc.scrollTop - top) > toc.clientHeight / 3) toc.scrollTop = top;
      }
    }
    var ticking = false;
    window.addEventListener("scroll", function () {
      if (ticking) return;
      ticking = true;
      requestAnimationFrame(function () { spy(); ticking = false; });
    }, { passive: true });
    spy();

    var toggle = $(".toc-toggle");
    if (toggle) {
      toggle.addEventListener("click", function () {
        var open = toc.classList.toggle("open");
        toggle.setAttribute("aria-expanded", String(open));
      });
      links.forEach(function (link) { link.addEventListener("click", function () { toc.classList.remove("open"); }); });
    }

    /* Filter: match the query against each section's full text, not only its title. */
    var search = $(".toc-search", toc);
    if (!search) return;
    var index = links.map(function (link, i) {
      var target = targets[i];
      var section = target && (target.tagName === "SECTION" ? target : target.closest("section"));
      var text = link.textContent;
      if (target && target.tagName === "H3") {
        var node = target.nextElementSibling;
        while (node && node.tagName !== "H3" && node.tagName !== "H2") { text += " " + node.textContent; node = node.nextElementSibling; }
      } else if (section) {
        text += " " + section.textContent;
      }
      return text.toLowerCase();
    });
    var none = $(".none", toc);
    search.addEventListener("input", function () {
      var words = search.value.toLowerCase().split(/\s+/).filter(Boolean);
      var shown = 0;
      links.forEach(function (link, i) {
        var hit = words.every(function (word) { return index[i].indexOf(word) !== -1; });
        link.parentNode.hidden = !hit;
        if (hit) shown++;
      });
      $$("h2", toc).forEach(function (group) {
        var list = group.nextElementSibling;
        group.hidden = !!list && $$("li:not([hidden])", list).length === 0;
      });
      if (none) none.style.display = shown ? "none" : "block";
    });
    document.addEventListener("keydown", function (event) {
      if (event.key === "/" && !/^(INPUT|TEXTAREA)$/.test(document.activeElement.tagName)) {
        event.preventDefault();
        toc.classList.add("open");
        search.focus();
      }
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    initTheme();
    initSetup();
    initCopy();
    initTabs();
    initTerminals();
    initLightbox();
    initGuide();
  });
})();
