// Простой движок презентаций: листание стрелками и кнопками, номер слайда в адресе (#3),
// подсветка синтаксиса JavaScript в блоках <pre class="code">.
(function () {
  const slides = Array.from(document.querySelectorAll('.slide'));
  if (!slides.length) return;

  const prev = document.querySelector('[data-prev]');
  const next = document.querySelector('[data-next]');
  const count = document.querySelector('[data-count]');
  const bar = document.querySelector('.progress');
  let current = 0;

  function show(i) {
    current = Math.max(0, Math.min(slides.length - 1, i));
    slides.forEach(function (s, idx) { s.classList.toggle('active', idx === current); });
    if (prev) prev.disabled = current === 0;
    if (next) next.disabled = current === slides.length - 1;
    if (count) count.textContent = (current + 1) + ' / ' + slides.length;
    if (bar) bar.style.width = ((current + 1) / slides.length * 100) + '%';
    if (location.hash !== '#' + (current + 1)) history.replaceState(null, '', '#' + (current + 1));
    window.scrollTo(0, 0);
  }

  if (prev) prev.addEventListener('click', function () { show(current - 1); });
  if (next) next.addEventListener('click', function () { show(current + 1); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'ArrowRight' || e.key === 'PageDown') show(current + 1);
    if (e.key === 'ArrowLeft' || e.key === 'PageUp') show(current - 1);
    if (e.key === 'Home') show(0);
    if (e.key === 'End') show(slides.length - 1);
  });

  const fromHash = parseInt(location.hash.slice(1), 10);
  show(isNaN(fromHash) ? 0 : fromHash - 1);

  // --- Подсветка синтаксиса ---
  const KEYWORDS = /^(const|let|var|function|return|if|else|for|of|in|while|new|true|false|null|undefined|throw|typeof|instanceof)$/;
  const esc = function (t) { return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  const TOKEN = /(\/\/[^\n]*|\/\*[\s\S]*?\*\/)|('(?:\\.|[^'\\])*'|"(?:\\.|[^"\\])*"|`(?:\\.|[^`\\])*`)|(\b\d+(?:\.\d+)?\b)|([A-Za-z_$][\w$]*)(\s*\()?|([\s\S])/g;

  document.querySelectorAll('pre.code').forEach(function (pre) {
    const src = pre.textContent;
    let out = '';
    src.replace(TOKEN, function (m, com, str, num, word, call, other) {
      if (com) out += '<span class="c">' + esc(com) + '</span>';
      else if (str) out += '<span class="s">' + esc(str) + '</span>';
      else if (num) out += '<span class="n">' + num + '</span>';
      else if (word) {
        if (KEYWORDS.test(word)) out += '<span class="k">' + word + '</span>' + (call ? esc(call) : '');
        else if (call) out += '<span class="f">' + word + '</span>' + esc(call);
        else out += word;
      } else out += esc(other);
      return m;
    });
    pre.innerHTML = out;
  });
})();
