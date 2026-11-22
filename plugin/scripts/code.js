(function (window, undefined) {

  var running = false;
  var ignoreUntil = 0;
  var timer = null;
  var skip = {};

  function $(id) { return document.getElementById(id); }

  function setStatus(msg, cls) {
    var s = $("status");
    if (!s) return;
    s.textContent = msg;
    s.className = cls || "";
  }


  function convert(manual) {
    if (running) return;
    running = true;
    if (manual) $("convertAll").disabled = true;
    window.Asc.scope.skip = skip;

    window.Asc.plugin.callCommand(function () {
      var doc = Api.GetDocument();
      if (typeof doc.AddMathEquation !== "function") {
        return { error: "AddMathEquation indisponible : ONLYOFFICE 8.2 ou plus requis." };
      }

      var skipList = Asc.scope.skip || {};
      var count = 0;
      var failed = [];


      var re = /\$\$([\s\S]+?)\$\$|(?<!\$)\$(?![\s$])([^$\n]*?[^$\s\\])\$(?![\d$])/g;

      var n = doc.GetElementsCount();
      for (var i = n - 1; i >= 0; i--) {
        var el = doc.GetElement(i);
        if (!el || el.GetClassType() !== "paragraph") continue;

        var text = el.GetText();
        if (!text || text.indexOf("$") === -1) continue;

        var seen = {}, found = [], m;
        re.lastIndex = 0;
        while ((m = re.exec(text)) !== null) {
          var raw = m[0];
          var latex = (m[1] !== undefined ? m[1] : m[2]).trim();
          if (!latex || seen[raw] || skipList[raw]) continue;
          seen[raw] = true;
          found.push({ raw: raw, latex: latex });
        }

        for (var f = 0; f < found.length; f++) {
          var guard = 0;
          while (guard++ < 50) {
            var ranges = el.Search(found[f].raw, true);
            if (!ranges || !ranges.length) break;
            ranges[ranges.length - 1].Select();
            if (!doc.AddMathEquation(found[f].latex, "latex")) { failed.push(found[f].raw); break; }
            count++;
            if (el.GetText().indexOf(found[f].raw) !== -1) { failed.push(found[f].raw); break; }
          }
        }
      }
      return { count: count, failed: failed };
    }, false, true, function (res) {
      if (res && res.failed) {
        for (var i = 0; i < res.failed.length; i++) skip[res.failed[i]] = true;
      }
      ignoreUntil = Date.now() + 1500;
      running = false;
      $("convertAll").disabled = false;

      if (!res) { setStatus("Aucune réponse de l'éditeur.", "err"); return; }
      if (res.error) { setStatus(res.error, "err"); return; }
      if (res.count > 0) setStatus(res.count + " équation(s) convertie(s).", "ok");
      else if (manual) setStatus("Aucune équation $...$ ou $$...$$ trouvée.");
      if (res.failed && res.failed.length) {
        setStatus((res.count || 0) + " convertie(s), " + res.failed.length + " ignorée(s).", "err");
      }
    });
  }

  function onCursorMoved() {
    if (!$("auto") || !$("auto").checked) return;
    if (running || Date.now() < ignoreUntil) return;
    clearTimeout(timer);
    timer = setTimeout(function () { convert(false); }, 400);
  }

  window.Asc.plugin.init = function () {
    $("convertAll").onclick = function () { skip = {}; convert(true); };
    $("auto").onchange = function () {
      setStatus(this.checked ? "Conversion automatique activée." : "Conversion automatique désactivée.");
      if (this.checked) convert(false);
    };

    if (typeof window.Asc.plugin.attachEditorEvent === "function") {
      window.Asc.plugin.attachEditorEvent("onTargetPositionChanged", onCursorMoved);
    }
  };


  window.Asc.plugin.event_onTargetPositionChanged = function () {
    if (typeof window.Asc.plugin.attachEditorEvent !== "function") onCursorMoved();
  };


  window.Asc.plugin.onThemeChanged = function (theme) {
    if (typeof window.Asc.plugin.onThemeChangedBase === "function") {
      window.Asc.plugin.onThemeChangedBase(theme);
    }
    try {
      document.body.style.color = theme["text-normal"] || "";
      document.body.style.background = theme["background-normal"] || "";
    } catch (e) {}
  };

  window.Asc.plugin.button = function () {};

})(window, undefined);
