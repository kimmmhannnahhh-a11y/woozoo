/* ===== 모바일 환수 검증기 — app logic ===== */
(function () {
  'use strict';

  // ---- 설정 ----
  const BASE = 'H';                 // 기준월 요금제그룹코드
  const COMPARE = ['K', 'N', 'Q'];  // M+1 · M+2 · M+3
  const SVC = 'D';                  // 서비스관리번호
  const DEVICE = 'E';               // 단말기
  const TERM = 'F';                 // 회선해지여부
  const HEADER_ROW = 1;
  const LS_KEY = 'woozoo_yogeum_acc_v1';
  const DL_HEAD = ['파일명', '행', '서비스관리번호', '단말기', '해지', '기준(H)', 'M+1(K)', 'M+2(N)', 'M+3(Q)'];

  // ---- 상태 ----
  let currentHits = [];     // 현재 업로드한 파일들의 해당 건 (각 항목에 file 포함)
  let accumulated = load();

  // ---- 유틸 ----
  function rankOf(code) { const m = (code || '').match(/\d+/); return m ? parseInt(m[0]) : NaN; }
  function colToIndex(c) { let n = 0; for (const ch of c) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function colName(n) { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; }
  function esc(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function escAttr(s) { return esc(s).replace(/"/g, '&quot;'); }
  function $(id) { return document.getElementById(id); }
  function load() { try { return JSON.parse(localStorage.getItem(LS_KEY) || '[]'); } catch (e) { return []; } }
  function save() { try { localStorage.setItem(LS_KEY, JSON.stringify(accumulated)); } catch (e) {} }
  function flash(el, text) { el.textContent = text; setTimeout(() => { el.textContent = ''; }, 4000); }

  // ---- 엑셀 셀 값 (지수표기 서비스관리번호 풀어서) ----
  function makeCellReader(shared) {
    return function (c) {
      const t = c.getAttribute('t');
      if (t === 'inlineStr') { const e = c.getElementsByTagName('t')[0]; return e ? e.textContent : ''; }
      const v = c.getElementsByTagName('v')[0];
      if (!v) return '';
      if (t === 's') return shared[parseInt(v.textContent)] || '';
      let raw = v.textContent;
      if (/^-?\d+(\.\d+)?[eE][+-]?\d+$/.test(raw)) {
        const n = Number(raw);
        if (!isNaN(n)) raw = n.toLocaleString('fullwide', { useGrouping: false });
      }
      return raw;
    };
  }

  // ---- 엑셀 1개 파싱 + 검증 → {fileTotal, hits} ----
  async function parseOne(file) {
    const fileName = file.name.replace(/\.xlsx$/i, '');
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const sheetPaths = Object.keys(zip.files).filter(p => /^xl\/worksheets\/sheet\d+\.xml$/.test(p)).sort();
    if (!sheetPaths.length) throw new Error('시트를 찾지 못했습니다');
    const xml = await zip.file(sheetPaths[0]).async('string');

    let shared = [];
    if (zip.file('xl/sharedStrings.xml')) {
      const ss = await zip.file('xl/sharedStrings.xml').async('string');
      const sdoc = new DOMParser().parseFromString(ss, 'application/xml');
      shared = [...sdoc.getElementsByTagName('si')].map(si => [...si.getElementsByTagName('t')].map(t => t.textContent).join(''));
    }
    const cellVal = makeCellReader(shared);
    const doc = new DOMParser().parseFromString(xml, 'application/xml');

    const want = [BASE, ...COMPARE, SVC, DEVICE, TERM];
    const widx = {}; want.forEach(c => widx[c] = colToIndex(c));

    const hits = [];
    let fileTotal = 0;
    [...doc.getElementsByTagName('row')].forEach(row => {
      const rn = parseInt(row.getAttribute('r'));
      if (rn <= HEADER_ROW) return;
      fileTotal++;
      const vals = {};
      [...row.getElementsByTagName('c')].forEach(c => {
        const ref = c.getAttribute('r'); const m = ref.match(/^([A-Z]+)/); if (!m) return;
        const idx = colToIndex(m[1]);
        for (const col in widx) if (idx === widx[col]) vals[col] = cellVal(c).trim();
      });
      const hRank = rankOf(vals[BASE] || '');
      let bad = false;
      for (const c of COMPARE) {
        const v = vals[c] || '';
        if (v === '') { bad = true; break; }                          // 빈값 → 해당
        if (!isNaN(hRank) && rankOf(v) < hRank) { bad = true; break; } // 등급 하향 → 해당
      }
      const isUsim = (vals[DEVICE] || '').toUpperCase().includes('USIM'); // USIM 은 무조건 제외
      const isTerm = (vals[TERM] || '').toUpperCase() === 'Y';            // 회선해지 Y
      if ((bad || isTerm) && !isUsim) {                                  // (요금제 사유 OR 해지 Y) & USIM 아님
        hits.push({
          file: fileName, rn, svc: vals[SVC] || '', device: vals[DEVICE] || '', term: vals[TERM] || '',
          h: vals[BASE] || '', vals: COMPARE.map(c => vals[c] || '')
        });
      }
    });
    return { fileTotal, hits };
  }

  // ---- 여러 파일 한 번에 처리 ----
  async function processFiles(fileList) {
    $('errBox').innerHTML = '';
    $('addMsg').textContent = '';
    const files = [...fileList].filter(f => f.name.toLowerCase().endsWith('.xlsx'));
    if (!files.length) { showErr('.xlsx 엑셀 파일만 업로드할 수 있습니다.'); return; }

    let total = 0; const allHits = []; const errs = [];
    for (const file of files) {
      try { const r = await parseOne(file); total += r.fileTotal; allHits.push(...r.hits); }
      catch (err) { errs.push(file.name + ': ' + err.message); console.error(err); }
    }
    currentHits = allHits;
    renderCurrent(total, allHits, files.length);
    if (errs.length) showErr('일부 파일 오류 — ' + errs.join(' / '));
  }

  // ---- 렌더: 검증 결과(현재) ----
  function renderCurrent(total, hits, fileCount) {
    $('statTotal').textContent = total.toLocaleString();
    $('statHit').textContent = hits.length.toLocaleString();
    $('statOk').textContent = (total - hits.length).toLocaleString();
    $('curCount').textContent = '(' + hits.length + '건' + (fileCount > 1 ? ' · ' + fileCount + '개 파일' : '') + ')';
    $('btnAdd').disabled = hits.length === 0;
    $('btnClear').disabled = total === 0;

    const tb = $('curBody');
    if (!hits.length) {
      tb.innerHTML = '<tr class="empty-row"><td colspan="10">' + (total ? '해당되는 건이 없습니다.' : '엑셀 파일을 업로드하면 결과가 표시됩니다.') + '</td></tr>';
      return;
    }
    tb.innerHTML = hits.map(h => {
      const hr = rankOf(h.h);
      return '<tr><td class="src">' + esc(h.file) + '</td><td>' + h.rn + '</td><td>' + cellOrBlank(h.svc) + '</td><td>' + cellOrBlank(h.device) + '</td><td>' + termCell(h.term) +
      '</td><td class="col-h">' + esc(h.h) + '</td><td>' + gradeCell(h.vals[0], hr) + '</td><td>' + gradeCell(h.vals[1], hr) +
      '</td><td>' + gradeCell(h.vals[2], hr) + '</td><td class="tag-hit">해당</td></tr>';
    }).join('');
  }

  // ---- 렌더: 모음 결과 ----
  function renderAcc() {
    $('accCount').textContent = '(' + accumulated.length + '건)';
    $('btnDownload').disabled = accumulated.length === 0;
    $('btnCopy').disabled = accumulated.length === 0;
    $('btnResetAll').disabled = accumulated.length === 0;
    const tb = $('accBody');
    if (!accumulated.length) {
      tb.innerHTML = '<tr class="empty-row"><td colspan="10">추가된 항목이 없습니다.</td></tr>';
      return;
    }
    tb.innerHTML = accumulated.map(a => {
      const hr = rankOf(a.h);
      return '<tr><td class="src">' + esc(a.file) + '</td><td>' + a.rn + '</td><td>' + cellOrBlank(a.svc) + '</td><td>' + cellOrBlank(a.device) +
      '</td><td>' + termCell(a.term) + '</td><td class="col-h">' + esc(a.h) + '</td><td>' + gradeCell(a.k, hr) + '</td><td>' + gradeCell(a.n, hr) +
      '</td><td>' + gradeCell(a.q, hr) + '</td><td><button class="icon-btn" data-del data-file="' + escAttr(a.file) + '" data-rn="' + a.rn + '" title="삭제">' + xIcon() + '</button></td></tr>';
    }).join('');
  }

  function cellOrBlank(v) { return (v == null || v === '') ? '<span class="cell-blank">(빈값)</span>' : esc(v); }
  // K·N·Q 셀: 빈값이거나 H보다 낮으면 빨간색
  function gradeCell(v, hRank) {
    if (v == null || v === '') return '<span class="cell-blank">(빈값)</span>';
    const r = rankOf(v);
    if (!isNaN(hRank) && !isNaN(r) && r < hRank) return '<span class="cell-down">' + esc(v) + '</span>';
    return esc(v);
  }
  function termCell(v) { return String(v).toUpperCase() === 'Y' ? '<span class="cell-term-y">Y</span>' : (v === '' || v == null ? '-' : esc(v)); }
  function xIcon() { return '<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M6 6l12 12M18 6L6 18"/></svg>'; }
  function showErr(msg) { $('errBox').innerHTML = '<div style="margin-top:16px;color:var(--danger);font-size:14px;">' + esc(msg) + '</div>'; }

  // ---- 새 xlsx 생성 ----
  async function buildXlsx(aoa, sheetName) {
    let rowsXml = '';
    aoa.forEach((row, ri) => {
      let cells = '';
      row.forEach((val, ci) => { cells += '<c r="' + colName(ci + 1) + (ri + 1) + '" t="inlineStr"><is><t xml:space="preserve">' + esc(val) + '</t></is></c>'; });
      rowsXml += '<row r="' + (ri + 1) + '">' + cells + '</row>';
    });
    const sheet = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' + rowsXml + '</sheetData></worksheet>';
    const ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>';
    const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>';
    const wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="' + esc(sheetName || 'Sheet1') + '" sheetId="1" r:id="rId1"/></sheets></workbook>';
    const wbr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>';
    const zip = new JSZip();
    zip.file('[Content_Types].xml', ct);
    zip.file('_rels/.rels', rels);
    zip.file('xl/workbook.xml', wb);
    zip.file('xl/_rels/workbook.xml.rels', wbr);
    zip.file('xl/worksheets/sheet1.xml', sheet);
    return await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function accToRows() { return accumulated.map(a => [a.file, a.rn, a.svc, a.device, a.term, a.h, a.k, a.n, a.q]); }

  // ---- 이벤트 ----
  function bind() {
    const drop = $('upload'), fileEl = $('fileInput');
    drop.onclick = () => fileEl.click();
    drop.ondragover = e => { e.preventDefault(); drop.classList.add('is-drag'); };
    drop.ondragleave = () => drop.classList.remove('is-drag');
    drop.ondrop = e => { e.preventDefault(); drop.classList.remove('is-drag'); if (e.dataTransfer.files.length) processFiles(e.dataTransfer.files); };
    fileEl.onchange = e => { if (e.target.files.length) { processFiles(e.target.files); fileEl.value = ''; } };

    // 업로드 데이터 추가 → 모음표 (각 건의 파일명 기준 중복 제외)
    $('btnAdd').onclick = () => {
      let added = 0;
      currentHits.forEach(h => {
        if (!accumulated.some(a => a.file === h.file && a.rn === h.rn)) {
          accumulated.push({ file: h.file, rn: h.rn, svc: h.svc, device: h.device, term: h.term, h: h.h, k: h.vals[0], n: h.vals[1], q: h.vals[2] });
          added++;
        }
      });
      save(); renderAcc();
      flash($('addMsg'), added > 0 ? added + '건 추가됨' : '이미 추가된 항목입니다');
    };

    // 초기화 (현재 검증 결과만)
    $('btnClear').onclick = () => { currentHits = []; renderCurrent(0, [], 0); };

    // 행별 삭제
    $('accBody').addEventListener('click', e => {
      const btn = e.target.closest('[data-del]'); if (!btn) return;
      const file = btn.getAttribute('data-file'); const rn = parseInt(btn.getAttribute('data-rn'));
      accumulated = accumulated.filter(a => !(a.file === file && a.rn === rn));
      save(); renderAcc();
    });

    // 엑셀 다운로드
    $('btnDownload').onclick = async () => {
      if (!accumulated.length) return;
      const blob = await buildXlsx([DL_HEAD].concat(accToRows()), '모음');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = '모바일환수검증_모음.xlsx';
      document.body.appendChild(a); a.click();
      setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 1500);
    };

    // 전체 복사
    $('btnCopy').onclick = async () => {
      if (!accumulated.length) return;
      const tsv = [DL_HEAD.join('\t')].concat(accToRows().map(r => r.join('\t'))).join('\n');
      try { await navigator.clipboard.writeText(tsv); flash($('accMsg'), '복사됨'); }
      catch (e) {
        const ta = document.createElement('textarea'); ta.value = tsv; document.body.appendChild(ta); ta.select();
        try { document.execCommand('copy'); flash($('accMsg'), '복사됨'); } catch (_) { flash($('accMsg'), '복사 실패'); }
        document.body.removeChild(ta);
      }
    };

    // 모음표 초기화
    $('btnResetAll').onclick = () => {
      if (!accumulated.length) return;
      if (!confirm('모음표 ' + accumulated.length + '건을 모두 삭제할까요?')) return;
      accumulated = []; save(); renderAcc();
    };

    // 자세히 보기 모달
    $('btnDetail').onclick = () => $('ruleModal').hidden = false;
    $('btnModalClose').onclick = () => $('ruleModal').hidden = true;
    $('ruleModal').addEventListener('click', e => { if (e.target === $('ruleModal')) $('ruleModal').hidden = true; });
  }

  // ---- 초기화 ----
  function init() {
    bind();
    renderCurrent(0, [], 0);
    renderAcc();
  }
  document.addEventListener('DOMContentLoaded', init);
})();
