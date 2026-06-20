(function () {
  'use strict';

  const TABLES = {
    cycles: 'eotq_cycles',
    nominees: 'eotq_nominees',
    questions: 'eotq_questions',
    responses: 'eotq_responses',
  };

  const QUESTION_TYPES = [
    { value: 'rating', label: 'Skala 1-5', scored: true },
    { value: 'single', label: 'Pilihan Tunggal', scored: true },
    { value: 'multi', label: 'Pilihan Ganda', scored: true },
    { value: 'text', label: 'Jawaban Teks', scored: false },
  ];

  function headers(extra) {
    return Object.assign({}, window.SUPABASE_HEADERS || {}, extra || {});
  }

  function apiUrl(path) {
    return `${SUPABASE_URL}/rest/v1/${path}`;
  }

  async function read(path) {
    const res = await fetch(apiUrl(path), { headers: headers() });
    if (!res.ok) throw await restError(res);
    return res.json();
  }

  async function write(path, method, payload, extraHeaders) {
    const res = await fetch(apiUrl(path), {
      method,
      headers: headers(Object.assign({ Prefer: 'return=representation' }, extraHeaders || {})),
      body: JSON.stringify(payload || {}),
    });
    if (!res.ok) throw await restError(res);
    const text = await res.text();
    return text ? JSON.parse(text) : null;
  }

  async function remove(path) {
    const res = await fetch(apiUrl(path), { method: 'DELETE', headers: headers({ Prefer: 'return=minimal' }) });
    if (!res.ok) throw await restError(res);
    return true;
  }

  async function restError(res) {
    let msg = `HTTP ${res.status}`;
    try {
      const err = await res.json();
      msg = err.message || err.details || err.hint || msg;
    } catch (_) {}
    return new Error(msg);
  }

  function qs(value) {
    return encodeURIComponent(String(value == null ? '' : value));
  }

  function number(value, fallback) {
    const n = Number(value);
    return Number.isFinite(n) ? n : (fallback || 0);
  }

  function fmtDateTime(value) {
    if (!value) return '-';
    const d = new Date(value);
    if (!Number.isFinite(d.getTime())) return '-';
    return d.toLocaleString('id-ID', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit',
    });
  }

  function toLocalInputValue(value) {
    if (!value) return '';
    const d = new Date(value);
    if (!Number.isFinite(d.getTime())) return '';
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  function fromLocalInputValue(value) {
    if (!value) return null;
    const d = new Date(value);
    return Number.isFinite(d.getTime()) ? d.toISOString() : null;
  }

  function cycleState(cycle, nowValue) {
    const now = nowValue ? new Date(nowValue) : new Date();
    const start = cycle && cycle.start_at ? new Date(cycle.start_at) : null;
    const end = cycle && cycle.end_at ? new Date(cycle.end_at) : null;
    const announce = cycle && cycle.announce_at ? new Date(cycle.announce_at) : end;
    const status = String(cycle && cycle.status || 'draft').toLowerCase();
    if (!cycle) return { key: 'empty', label: 'Belum Ada', open: false, results: false };
    if (status === 'archived') return { key: 'archived', label: 'Arsip', open: false, results: false };
    if (status !== 'published') return { key: 'draft', label: 'Draft', open: false, results: false };
    if (!start || !end) return { key: 'draft', label: 'Belum Mulai', open: false, results: false };
    if (start && now < start) return { key: 'draft', label: 'Belum Mulai', open: false, results: false };
    if (start && end && now >= start && now <= end) return { key: 'open', label: 'Penilaian Aktif', open: true, results: false };
    if (announce && now < announce) return { key: 'closed', label: 'Menunggu Pengumuman', open: false, results: false };
    return { key: 'announced', label: 'Pengumuman', open: false, results: true };
  }

  function parseOptions(text) {
    return String(text || '')
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(Boolean)
      .map((line, idx) => {
        const parts = line.split('|').map(x => x.trim());
        const label = parts[0] || `Opsi ${idx + 1}`;
        const score = parts.length > 1 ? number(parts[1], 0) : 0;
        return { label, score };
      });
  }

  function optionsToText(options) {
    return (Array.isArray(options) ? options : [])
      .map(opt => `${opt.label || ''}|${number(opt.score, 0)}`)
      .join('\n');
  }

  function answerScore(question, value) {
    const weight = number(question && question.weight, 1);
    const type = question && question.type;
    if (type === 'rating') return number(value, 0) * weight;
    if (type === 'single') {
      const opt = (question.options || []).find(o => String(o.label) === String(value));
      return number(opt && opt.score, 0) * weight;
    }
    if (type === 'multi') {
      const selected = Array.isArray(value) ? value : [];
      return selected.reduce((sum, label) => {
        const opt = (question.options || []).find(o => String(o.label) === String(label));
        return sum + number(opt && opt.score, 0);
      }, 0) * weight;
    }
    return 0;
  }

  function buildAnswers(questions, formValues) {
    return (questions || []).map(q => {
      const value = formValues[String(q.id)];
      const score = answerScore(q, value);
      return {
        question_id: q.id,
        question: q.question,
        type: q.type,
        weight: number(q.weight, 1),
        value,
        score,
      };
    });
  }

  function totalScore(answers) {
    return (answers || []).reduce((sum, a) => sum + number(a && a.score, 0), 0);
  }

  function summarize(nominees, responses) {
    const map = new Map((nominees || []).map(n => [String(n.id), {
      nominee: n,
      total: 0,
      votes: 0,
      average: 0,
    }]));
    (responses || []).forEach(r => {
      const item = map.get(String(r.nominee_id));
      if (!item) return;
      item.total += number(r.total_score, 0);
      item.votes += 1;
    });
    return Array.from(map.values())
      .map(item => Object.assign(item, { average: item.votes ? item.total / item.votes : 0 }))
      .sort((a, b) => (b.average - a.average) || (b.total - a.total) || String(a.nominee.pegawai_nama || '').localeCompare(String(b.nominee.pegawai_nama || '')));
  }

  async function loadCurrentPegawai(session) {
    if (!session || !session.id) return null;
    let rows = await read(`data_pegawai?id=eq.${qs(session.id)}&select=*&limit=1`).catch(() => []);
    if (rows && rows[0]) return rows[0];
    rows = await read(`data_pegawai?pegawai_nip=eq.${qs(session.username)}&select=*&limit=1`).catch(() => []);
    return rows && rows[0] ? rows[0] : null;
  }

  const Eotq = {
    QUESTION_TYPES,
    fmtDateTime,
    toLocalInputValue,
    fromLocalInputValue,
    cycleState,
    parseOptions,
    optionsToText,
    answerScore,
    buildAnswers,
    totalScore,
    summarize,
    loadCurrentPegawai,
    loadEmployees: () => read('data_pegawai?select=*&order=nama.asc'),
    loadCycles: () => read(`${TABLES.cycles}?select=*&order=created_at.desc`),
    loadCycleBundle: async cycleId => {
      const id = qs(cycleId);
      const [cycleRows, nominees, questions, responses] = await Promise.all([
        read(`${TABLES.cycles}?id=eq.${id}&select=*&limit=1`),
        read(`${TABLES.nominees}?cycle_id=eq.${id}&select=*&order=sort_order.asc,created_at.asc`),
        read(`${TABLES.questions}?cycle_id=eq.${id}&select=*&order=sort_order.asc,created_at.asc`),
        read(`${TABLES.responses}?cycle_id=eq.${id}&select=*&order=created_at.asc`),
      ]);
      return { cycle: cycleRows[0] || null, nominees, questions, responses };
    },
    createCycle: payload => write(TABLES.cycles, 'POST', payload),
    updateCycle: (id, payload) => write(`${TABLES.cycles}?id=eq.${qs(id)}`, 'PATCH', payload),
    deleteCycle: id => remove(`${TABLES.cycles}?id=eq.${qs(id)}`),
    addNominee: payload => write(TABLES.nominees, 'POST', payload),
    updateNominee: (id, payload) => write(`${TABLES.nominees}?id=eq.${qs(id)}`, 'PATCH', payload),
    deleteNominee: id => remove(`${TABLES.nominees}?id=eq.${qs(id)}`),
    addQuestion: payload => write(TABLES.questions, 'POST', payload),
    updateQuestion: (id, payload) => write(`${TABLES.questions}?id=eq.${qs(id)}`, 'PATCH', payload),
    deleteQuestion: id => remove(`${TABLES.questions}?id=eq.${qs(id)}`),
    submitResponse: payload => write(
      `${TABLES.responses}?on_conflict=cycle_id,nominee_id,voter_user_id`,
      'POST',
      payload,
      { Prefer: 'resolution=merge-duplicates,return=representation' }
    ),
  };

  window.Eotq9201 = Eotq;
})();
