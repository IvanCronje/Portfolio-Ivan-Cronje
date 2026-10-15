/* Shared site navigation. Add <div id="site-nav"></div> at the top of <body>.
   Optional: data-parent="library" for pages opened from the Working library (adds a breadcrumb),
   data-title to override the breadcrumb title. Links live in one place: edit LINKS below. */
(function () {
  var LINKS = [
    { k: "home", t: "Portfolio", h: "index.html" },
    { k: "systems", t: "Systems", h: "systems.html" },
    { k: "lab", t: "Lab", h: "lab.html" },
    { k: "library", t: "Library", h: "library.html" },
    { k: "stories", t: "Stories", h: "stories.html" },
    { k: "vision", t: "Vision", h: "vision.html" },
    { k: "philosophy", t: "Philosophy", h: "philosophy.html" }
  ];
  var host = document.getElementById("site-nav");
  if (!host) return;
  var file = (location.pathname.split("/").pop() || "index.html").replace(".html", "");
  var parent = host.getAttribute("data-parent");
  var active = parent || (file === "index" ? "home" : file);
  var css = "#site-nav{position:sticky;top:0;z-index:999;font-family:ui-monospace,SFMono-Regular,Menlo,monospace}" +
    "#site-nav .bar{display:flex;align-items:center;justify-content:space-between;gap:20px;padding:0 28px;height:56px;background:rgba(11,12,13,.9);backdrop-filter:blur(12px);-webkit-backdrop-filter:blur(12px);border-bottom:1px solid rgba(242,242,243,.16)}" +
    "#site-nav .logo{font-family:var(--font-heading,Bahnschrift,'Arial Narrow',sans-serif);font-size:19px;letter-spacing:.07em;color:#f2f2f3;text-decoration:none;white-space:nowrap;flex-shrink:0}" +
    "#site-nav .links{display:flex;align-items:center;gap:6px;overflow-x:auto;scrollbar-width:none;min-width:0}#site-nav .links::-webkit-scrollbar{display:none}" +
    "#site-nav .links a{font-size:10.5px;letter-spacing:.15em;text-transform:uppercase;color:rgba(242,242,243,.66);text-decoration:none;padding:8px 10px;white-space:nowrap;border-bottom:1px solid transparent}" +
    "#site-nav .links a:hover{color:#8fb3d6}#site-nav .links a.on{color:#f2f2f3;border-bottom-color:#5980a6}" +
    "#site-nav .cta{font-size:10.5px;letter-spacing:.14em;text-transform:uppercase;color:#0b0c0d;background:#5980a6;padding:9px 14px;text-decoration:none;white-space:nowrap;flex-shrink:0}" +
    "#site-nav .crumb{padding:9px 28px;background:#101214;border-bottom:1px solid rgba(242,242,243,.12);font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:rgba(242,242,243,.55);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}" +
    "#site-nav .crumb a{color:#8fb3d6;text-decoration:none}#site-nav .crumb a:hover{color:#f2f2f3}#site-nav .crumb i{font-style:normal;margin:0 8px;opacity:.5}" +
    "@media(max-width:760px){#site-nav .bar{padding:0 14px;gap:10px}#site-nav .cta{display:none}#site-nav .crumb{padding:9px 14px}}";
  var s = document.createElement("style"); s.textContent = css; document.head.appendChild(s);
  var links = LINKS.map(function (l) {
    return '<a href="' + l.h + '"' + (l.k === active ? ' class="on" aria-current="page"' : "") + ">" + l.t + "</a>";
  }).join("");
  var html = '<nav class="bar" aria-label="Site"><a class="logo" href="index.html">IVAN CRONJE</a><div class="links">' + links +
    '</div><a class="cta" href="index.html#contact">Get in touch</a></nav>';
  var PAR = { library: ["Working library", "library.html"], lab: ["Lab", "lab.html"] };
  if (PAR[parent]) {
    var title = host.getAttribute("data-title") || document.title.split(/\s[—|-]\s/)[0];
    html += '<div class="crumb"><a href="index.html">Portfolio</a><i>/</i><a href="' + PAR[parent][1] + '">' + PAR[parent][0] + '</a><i>/</i>' +
      title.replace(/</g, "&lt;") + "</div>";
  }
  host.innerHTML = html;
  var on = host.querySelector(".links a.on"), row = host.querySelector(".links");
  if (on && row) row.scrollLeft = Math.max(0, on.offsetLeft - 24);
})();
