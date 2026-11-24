// ===== بخش پاپ‌آپ =====
if (typeof document !== 'undefined' && document.getElementById('extractBtn')) {
  const btn = document.getElementById('extractBtn');
  const statusEl = document.getElementById('status');

  const setStatus = (msg, type = '') => {
    statusEl.textContent = msg;
    statusEl.className = type;
  };

  // ذخیره و بازیابی تنظیمات کاربر
  const OPT_IDS = ['optUrl', 'optTags', 'optLinks', 'optImages'];
  if (chrome.storage && chrome.storage.local) {
    chrome.storage.local.get(OPT_IDS, saved => {
      OPT_IDS.forEach(id => { if (typeof saved[id] === 'boolean') document.getElementById(id).checked = saved[id]; });
    });
    OPT_IDS.forEach(id => document.getElementById(id).addEventListener('change', e => {
      chrome.storage.local.set({ [id]: e.target.checked });
    }));
  }

  // لینک کانال تلگرام در تب جدید
  document.getElementById('authorLink').addEventListener('click', e => {
    e.preventDefault();
    chrome.tabs.create({ url: 'https://t.me/majidvosoughy' });
  });

  btn.addEventListener('click', async () => {
    btn.disabled = true;
    setStatus('در حال استخراج...');

    try {
      const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
      if (!tab || !tab.id || /^(chrome|edge|about|chrome-extension):/.test(tab.url || '')) {
        throw new Error('این صفحه قابل استخراج نیست.');
      }

      const opts = {
        images: document.getElementById('optImages').checked,
        links: document.getElementById('optLinks').checked
      };

      // تزریق کتابخانه Readability (برای سایت‌های غیروردپرسی)
      try {
        await chrome.scripting.executeScript({ target: { tabId: tab.id }, files: ['Readability.js'] });
      } catch (e) { /* اگر نشد، روش‌های دیگر استفاده می‌شوند */ }

      const results = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: extractArticle,
        args: [opts]
      });

      const data = results && results[0] && results[0].result;
      if (!data || !data.html) throw new Error('محتوایی پیدا نشد.');

      downloadWord(data, {
        includeUrl: document.getElementById('optUrl').checked,
        showTags: document.getElementById('optTags').checked
      });
      setStatus(`انجام شد: ${data.headings} هدینگ، ${data.paragraphs} پاراگراف، حدود ${data.words} کلمه`, 'ok');
    } catch (e) {
      setStatus('خطا: ' + (e.message || e), 'err');
    } finally {
      btn.disabled = false;
    }
  });
}

// ساخت فایل Word و دانلود از داخل پاپ‌آپ
function downloadWord(data, wordOpts) {
  wordOpts = wordOpts || {};
  const dir = data.dir === 'ltr' ? 'ltr' : 'rtl';
  const align = dir === 'rtl' ? 'right' : 'left';
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

  let html = data.html;

  // نمایش نوع تگ قبل از متن هر هدینگ: [H2] عنوان
  if (wordOpts.showTags) {
    html = html.replace(/<(h[1-6])>/g, (m, t) =>
      '<' + t + '><span style="color:#2563eb;font-family:Arial">[' + t.toUpperCase() + ']</span> ');
  }

  // درج آدرس صفحه در ابتدای فایل
  if (wordOpts.includeUrl && data.url) {
    const u = esc(data.url).replace(/"/g, '&quot;');
    html = '<p style="direction:ltr;text-align:left;font-size:10pt;color:#555">' +
           'URL: <a href="' + u + '">' + esc(decodeURIComponentSafe(data.url)) + '</a></p>' + html;
  }

  const doc =
    "<html xmlns:o='urn:schemas-microsoft-com:office:office' " +
    "xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>" +
    "<head><meta charset='utf-8'><title>" + esc(data.title) + "</title>" +
    "<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>100</w:Zoom></w:WordDocument></xml><![endif]-->" +
    "<style>" +
    "body{font-family:Tahoma,Arial,sans-serif;font-size:12pt;line-height:1.7;direction:" + dir + ";text-align:" + align + ";}" +
    "p,li,h1,h2,h3,h4,h5,h6,td,th,blockquote{direction:" + dir + ";text-align:" + align + ";}" +
    "h1{font-size:20pt;} h2{font-size:16pt;} h3{font-size:14pt;} h4,h5,h6{font-size:12pt;}" +
    "table{border-collapse:collapse;} td,th{border:1px solid #999;padding:4px 8px;}" +
    "img{max-width:600px;}" +
    "</style></head>" +
    "<body dir='" + dir + "' lang='fa'>" + html + "</body></html>";

  const blob = new Blob(['﻿', doc], { type: 'application/msword' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = (data.fileName || 'Article') + '.doc';
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}

function decodeURIComponentSafe(u) {
  try { return decodeURI(u); } catch (e) { return u; }
}

// ===== این تابع داخل صفحه وب اجرا می‌شود (باید کاملاً مستقل باشد) =====
function extractArticle(opts) {
  opts = opts || {};

  // ---------- ۱. پیدا کردن بدنه اصلی مقاله ----------
  // سلکتورهای اختصاصی CMSهای رایج (وردپرس، المنتور، جوملا، دروپال، Ghost، بلاگر، ...)
  const CONTENT_SELECTORS = [
    '.elementor-widget-theme-post-content',
    '[itemprop="articleBody"]',
    '.entry-content', '.post-content', '.article-content', '.article-body',
    '.post-body', '.single-post-content', '.single-content', '.td-post-content',
    '.post-entry', '.blog-content', '.content-body', '.news-body', '.news-text',
    '.item-page', '[itemprop="text"]', '.field--name-body', '.node__content',
    '.gh-content', '.post-full-content', '.post-body-inner', '.story-body',
    '.article__body', '.article-text', '.content-text', '.markdown-body', '.prose'
  ];

  // طول متن «مقاله‌ای» یک عنصر: پاراگراف، لیست، هدینگ، سلول جدول
  const textLen = el => {
    let n = 0;
    el.querySelectorAll('p, li, h2, h3, h4, td, th, pre, blockquote').forEach(x => {
      if (!x.parentElement.closest('p, li, td, th, blockquote')) n += x.textContent.trim().length;
    });
    return n;
  };

  function bySelectors() {
    for (const sel of CONTENT_SELECTORS) {
      let best = null, bestLen = 0;
      document.querySelectorAll(sel).forEach(el => {
        const len = textLen(el);
        if (len > bestLen) { best = el; bestLen = len; }
      });
      if (best && bestLen > 400) return best;
    }
    return null;
  }

  // روش عمومی برای هر سایتی: کتابخانه Readability فایرفاکس
  // جدول‌ها قبل از پردازش با علامت جایگزین می‌شوند تا حتماً حفظ شوند
  function byReadability() {
    if (typeof Readability === 'undefined') return null;
    try {
      const docClone = document.cloneNode(true);
      const tables = [];
      docClone.querySelectorAll('table').forEach(t => {
        if (t.parentElement && t.parentElement.closest('table')) return; // جدول تودرتو
        if (!t.textContent.trim()) return;
        const i = tables.length;
        tables.push(t.cloneNode(true));
        const ph = docClone.createElement('p');
        ph.textContent = '[[XTABLE' + i + ']]';
        t.replaceWith(ph);
      });

      const art = new Readability(docClone, { charThreshold: 300 }).parse();
      if (!art || !art.content) return null;

      const root = new DOMParser().parseFromString(art.content, 'text/html').body;
      tables.forEach((tbl, i) => {
        const marker = '[[XTABLE' + i + ']]';
        const walker = root.ownerDocument.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        let tn;
        while ((tn = walker.nextNode())) {
          if (tn.nodeValue.includes(marker)) {
            let el = tn.parentElement;
            while (el.parentElement && el.parentElement !== root &&
                   el.parentElement.textContent.trim() === marker) el = el.parentElement;
            el.replaceWith(root.ownerDocument.importNode(tbl, true));
            break;
          }
        }
      });
      return { node: root, title: art.title };
    } catch (e) {
      return null;
    }
  }

  // آخرین راه: عنصری که بیشترین متن (پاراگراف/جدول/لیست) را دارد
  function byScoring() {
    const scores = new Map();
    const add = (node, len) => {
      let el = node.parentElement, w = 1;
      for (let i = 0; i < 3 && el; i++) {
        scores.set(el, (scores.get(el) || 0) + len * w);
        w /= 2;
        el = el.parentElement;
      }
    };
    document.querySelectorAll('p, pre, blockquote, table, ul, ol, div').forEach(x => {
      if (x.closest('nav, header, footer, aside, form')) return;
      let len;
      if (x.tagName === 'DIV') {
        // متن مستقیم داخل div (سایت‌هایی که از p استفاده نمی‌کنند)
        len = 0;
        x.childNodes.forEach(c => { if (c.nodeType === 3) len += c.nodeValue.trim().length; });
      } else {
        len = x.textContent.trim().length;
        const links = [...x.querySelectorAll('a')].reduce((n, a) => n + a.textContent.length, 0);
        if (len && links / len > 0.5) return; // منو یا لیست لینک
      }
      if (len >= 25) add(x, len);
    });
    let best = document.body, bestScore = 0;
    scores.forEach((s, el) => { if (s > bestScore) { bestScore = s; best = el; } });
    return best;
  }

  let contentNode = bySelectors();
  let readTitle = '';
  const rd = byReadability();
  if (rd) readTitle = rd.title || '';
  // اگر سلکتور پیدا نشد، یا Readability متن خیلی بیشتری پیدا کرد، از Readability استفاده کن
  if (rd && (!contentNode || textLen(rd.node) > textLen(contentNode) * 1.5)) contentNode = rd.node;
  if (!contentNode || textLen(contentNode) < 200) {
    const sc = byScoring();
    if (!contentNode || textLen(sc) > textLen(contentNode)) contentNode = sc;
  }

  // ---------- ۲. عنوان مقاله ----------
  const h1 = document.querySelector(
    '.elementor-widget-theme-post-title h1, h1.entry-title, h1.post-title, article h1, main h1, h1'
  );
  const ogTitle = document.querySelector('meta[property="og:title"]');
  const title = (h1 && h1.textContent.trim()) ||
                (ogTitle && ogTitle.content.trim()) || readTitle ||
                document.title.trim() || 'Article';

  // ---------- ۳. پاک‌سازی ----------
  const clone = contentNode.cloneNode(true);
  const ALWAYS_REMOVE = 'script, style, noscript, iframe, form, button, input, select, textarea, svg, canvas, video, audio, template, [hidden]';
  const JUNK = [
    'nav', 'aside', 'footer', '.sidebar', '.widget', '.comments', '#comments', '.ads', '.ad',
    '[class*="advert"]', '.sharedaddy', '[class*="share-"]', '[class*="social"]', '[class*="related"]',
    '[class*="breadcrumb"]', '.ez-toc-container', '#ez-toc-container', '.lwptoc', '#toc', '.toc',
    '.elementor-widget-table-of-contents', '.mwe-toc', '[class*="mwe-inline"]',
    '.screen-reader-text', '.sr-only'
  ];
  clone.querySelectorAll(ALWAYS_REMOVE).forEach(el => el.remove());
  // بخش‌های نامربوط حذف می‌شوند، مگر اینکه جدول داخلشان باشد
  clone.querySelectorAll(JUNK.join(',')).forEach(el => {
    if (!el.querySelector('table')) el.remove();
  });

  // ---------- ۴. ساخت HTML تمیز ----------
  const HEADINGS = new Set(['H1', 'H2', 'H3', 'H4', 'H5', 'H6']);
  const TEXT_BLOCKS = new Set(['P', 'LI', 'TD', 'TH', 'BLOCKQUOTE', 'PRE', 'FIGCAPTION', 'CAPTION', 'DT', 'DD']);
  const STRUCT = new Set(['UL', 'OL', 'TABLE', 'THEAD', 'TBODY', 'TFOOT', 'TR', 'DL']);
  const INLINE_KEEP = { STRONG: 'b', B: 'b', EM: 'i', I: 'i', U: 'u', SUP: 'sup', SUB: 'sub', CODE: 'code', MARK: 'b', S: 's', DEL: 's' };
  const BLOCKISH = new Set([...HEADINGS, ...TEXT_BLOCKS, ...STRUCT,
    'DIV', 'SECTION', 'ARTICLE', 'MAIN', 'FIGURE', 'HEADER', 'HR', 'ADDRESS', 'DETAILS', 'SUMMARY', 'CENTER', 'IMG']);

  const esc = s => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const escAttr = s => esc(s).replace(/"/g, '&quot;');
  const abs = u => { try { return new URL(u, location.href).href; } catch (e) { return ''; } };
  const hasText = html =>
    /<img/i.test(html) ||
    html.replace(/<[^>]+>/g, '').replace(/&nbsp;|[\s‌‏‎]/g, '') !== '';

  function imgHtml(img) {
    if (!opts.images) return '';
    let src = img.getAttribute('data-src') || img.getAttribute('data-lazy-src') ||
              img.getAttribute('data-original') || '';
    if (!src) {
      const ss = img.getAttribute('data-srcset') || img.getAttribute('srcset') || '';
      if (ss) src = ss.split(',').pop().trim().split(/\s+/)[0];
    }
    if (!src) src = img.getAttribute('src') || '';
    if (!src || src.startsWith('data:')) return '';
    const alt = img.getAttribute('alt') || '';
    return '<p style="text-align:center"><img src="' + escAttr(abs(src)) + '" alt="' + escAttr(alt) + '" width="550"></p>';
  }

  // حالت mixed: wrapInline=true یعنی متن‌های پراکنده در div را داخل <p> بگذار
  function serChildren(node, wrapInline) {
    let out = '', buf = '';
    const flush = () => { if (hasText(buf)) out += '<p>' + buf.trim() + '</p>'; buf = ''; };
    node.childNodes.forEach(c => {
      const isBlock = c.nodeType === 1 && BLOCKISH.has(c.tagName);
      const html = ser(c);
      if (wrapInline && isBlock) { flush(); out += html; }
      else if (wrapInline) buf += html;
      else out += html;
    });
    if (wrapInline) flush();
    return out;
  }

  function ser(node) {
    if (node.nodeType === 3) return esc(node.nodeValue.replace(/[\t\n\r ]+/g, ' '));
    if (node.nodeType !== 1) return '';
    const tag = node.tagName;

    if (tag === 'BR') return '<br>';
    if (tag === 'IMG') return imgHtml(node);
    if (tag === 'HR') return '';

    if (HEADINGS.has(tag)) {
      const inner = serChildren(node, false).replace(/<img[^>]*>/g, '').trim();
      return hasText(inner) ? '<' + tag.toLowerCase() + '>' + inner + '</' + tag.toLowerCase() + '>' : '';
    }
    if (TEXT_BLOCKS.has(tag)) {
      const inner = serChildren(node, false).trim();
      if (!hasText(inner)) return tag === 'TD' || tag === 'TH' ? '<' + tag.toLowerCase() + '></' + tag.toLowerCase() + '>' : '';
      let attrs = '';
      if (tag === 'TD' || tag === 'TH') {
        ['colspan', 'rowspan'].forEach(a => { if (node.getAttribute(a)) attrs += ' ' + a + '="' + escAttr(node.getAttribute(a)) + '"'; });
      }
      const t = tag === 'FIGCAPTION' || tag === 'CAPTION' ? 'p' : tag.toLowerCase();
      return '<' + t + attrs + '>' + inner + '</' + t + '>';
    }
    if (STRUCT.has(tag)) {
      const inner = serChildren(node, false);
      if (!hasText(inner)) return '';
      if (tag === 'TABLE') return '<table border="1" cellpadding="5" style="border-collapse:collapse">' + inner + '</table>';
      return '<' + tag.toLowerCase() + '>' + inner + '</' + tag.toLowerCase() + '>';
    }
    if (INLINE_KEEP[tag]) {
      const inner = serChildren(node, false);
      if (!hasText(inner)) return inner;
      const t = INLINE_KEEP[tag];
      return '<' + t + '>' + inner + '</' + t + '>';
    }
    if (tag === 'A') {
      const inner = serChildren(node, false);
      const href = node.getAttribute('href') || '';
      if (opts.links && href && !href.startsWith('#') && !/^javascript:/i.test(href) && hasText(inner)) {
        return '<a href="' + escAttr(abs(href)) + '">' + inner + '</a>';
      }
      return inner;
    }
    // سایر تگ‌ها (div, section, span, figure, ...) باز می‌شوند
    return serChildren(node, BLOCKISH.has(tag));
  }

  let body = serChildren(clone, true);

  // عنوان همیشه به صورت H1 در ابتدای فایل؛ تکرار عنوان در اول متن حذف می‌شود
  const norm = t => t.replace(/<[^>]+>/g, '').replace(/&amp;/g, '&').replace(/[\s\u200c]+/g, ' ').trim();
  const firstH = body.match(/^\s*<(h[1-3])>([\s\S]*?)<\/\1>/);
  if (firstH && norm(firstH[2]) === norm(esc(title))) body = body.slice(firstH[0].length);
  body = '<h1>' + esc(title) + '</h1>' + body;

  // ---------- ۵. آمار و خروجی ----------
  const plain = body.replace(/<[^>]+>/g, ' ');
  const dir = (document.documentElement.getAttribute('dir') ||
               (document.body && getComputedStyle(document.body).direction) || 'rtl').toLowerCase();

  return {
    title: title,
    url: location.href,
    fileName: title.replace(/[\/\\?%*:|"<>]/g, '-').slice(0, 120),
    html: body,
    dir: dir === 'ltr' ? 'ltr' : 'rtl',
    headings: (body.match(/<h[1-6][\s>]/g) || []).length,
    paragraphs: (body.match(/<p[\s>]/g) || []).length,
    words: plain.split(/\s+/).filter(Boolean).length
  };
}

if (typeof module !== 'undefined') module.exports = { extractArticle };
