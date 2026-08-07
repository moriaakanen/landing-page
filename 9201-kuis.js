/**
 * 9201 KUIS ENGINE & PERSISTENCE
 * ─────────────────────────────────────────────────────────
 * Modul terpusat pengelola Kuis Interaktif (Kahoot/Wayground Style)
 * Mendukung Supabase API dengan fallback otomatis ke localStorage.
 * Dilengkapi Web Audio API Synthesizer (tanpa file audio eksternal).
 */

(function (global) {
  'use strict';

  const STORAGE_KEY = '9201_quizzes_v1';
  const ATTEMPTS_KEY = '9201_quiz_attempts_v1';

  // ─── SEED INITIAL LOCAL QUIZZES ───────────────────────────
  const INITIAL_QUIZZES = [
    {
      id: 'quiz-demo-9201',
      title: 'Kuis Pengetahuan Umum Portal 9201',
      description: 'Uji wawasan Anda seputar fitur, kepegawaian, dan budaya kerja Portal 9201 secara interaktif!',
      category: 'Kepegawaian',
      pin_code: '920101',
      is_active: true,
      time_per_question_sec: 20,
      created_by: 'Admin',
      created_at: new Date().toISOString(),
      questions: [
        {
          id: 'q1',
          order_seq: 1,
          question_text: 'Apa kepanjangan dari fitur EoTQ pada Portal 9201?',
          option_a: 'Employee of The Quarter',
          option_b: 'Evaluation of The Quality',
          option_c: 'Executive of The Quarter',
          option_d: 'Employee of The Quantity',
          correct_option: 'A',
          points: 100
        },
        {
          id: 'q2',
          order_seq: 2,
          question_text: 'Dokumen apa yang biasanya diajukan sebelum melakukan perjalanan dinas?',
          option_a: 'Kamus POK',
          option_b: 'Surat Tugas',
          option_c: 'Nota Beban Work',
          option_d: 'Sertifikat Pelatihan',
          correct_option: 'B',
          points: 100
        },
        {
          id: 'q3',
          order_seq: 3,
          question_text: 'Kombinasi warna utama yang mendominasi identitas visual Portal 9201 adalah...',
          option_a: 'Merah & Hitam',
          option_b: 'Hijau & Perak',
          option_c: 'Navy (#0d2340) & Gold (#c8a84b)',
          option_d: 'Kuning & Ungu',
          correct_option: 'C',
          points: 100
        },
        {
          id: 'q4',
          order_seq: 4,
          question_text: 'Menu apa yang digunakan admin untuk mengelola daftar pengguna dan role di portal?',
          option_a: 'Predikat Kinerja',
          option_b: 'Manajemen Pengguna',
          option_c: 'Buku Tamu',
          option_d: 'Manajemen Mitra',
          correct_option: 'B',
          points: 100
        }
      ]
    }
  ];

  // ─── WEB AUDIO API SYNTHESIZER ────────────────────────────
  let audioCtx = null;
  function getAudioContext() {
    if (!audioCtx) {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (AudioCtx) audioCtx = new AudioCtx();
    }
    if (audioCtx && audioCtx.state === 'suspended') {
      audioCtx.resume();
    }
    return audioCtx;
  }

  const KuisAudio = {
    playTick() {
      try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sine';
        osc.frequency.setValueAtTime(600, ctx.currentTime);
        gain.gain.setValueAtTime(0.08, ctx.currentTime);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.08);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start();
        osc.stop(ctx.currentTime + 0.08);
      } catch (e) {}
    },
    playCorrect() {
      try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const notes = [523.25, 659.25, 783.99, 1046.50]; // C5, E5, G5, C6
        notes.forEach((freq, idx) => {
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();
          osc.type = 'triangle';
          osc.frequency.setValueAtTime(freq, now + idx * 0.08);
          gain.gain.setValueAtTime(0.2, now + idx * 0.08);
          gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.08 + 0.25);
          osc.connect(gain);
          gain.connect(ctx.destination);
          osc.start(now + idx * 0.08);
          osc.stop(now + idx * 0.08 + 0.25);
        });
      } catch (e) {}
    },
    playWrong() {
      try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = 'sawtooth';
        osc.frequency.setValueAtTime(180, now);
        osc.frequency.linearRampToValueAtTime(110, now + 0.3);
        gain.gain.setValueAtTime(0.25, now);
        gain.gain.exponentialRampToValueAtTime(0.001, now + 0.3);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(now);
        osc.stop(now + 0.3);
      } catch (e) {}
    },
    playFanfare() {
      try {
        const ctx = getAudioContext();
        if (!ctx) return;
        const now = ctx.currentTime;
        const arpeggio = [523.25, 659.25, 783.99, 1046.50, 1318.51, 1567.98];
        arpeggio.forEach((freq, i) => {
          const osc = ctx.createOscillator();
          const gain = ctx.createGain();
          osc.type = 'sine';
          osc.frequency.setValueAtTime(freq, now + i * 0.1);
          gain.gain.setValueAtTime(0.25, now + i * 0.1);
          gain.gain.exponentialRampToValueAtTime(0.001, now + i * 0.1 + 0.4);
          osc.connect(gain);
          gain.connect(ctx.destination);
          osc.start(now + i * 0.1);
          osc.stop(now + i * 0.1 + 0.4);
        });
      } catch (e) {}
    }
  };

  // ─── LOCAL STORAGE HELPERS ────────────────────────────────
  function loadLocalQuizzes() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(INITIAL_QUIZZES));
        return INITIAL_QUIZZES;
      }
      return JSON.parse(raw);
    } catch (e) {
      return INITIAL_QUIZZES;
    }
  }

  function saveLocalQuizzes(quizzes) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(quizzes));
  }

  function loadLocalAttempts() {
    try {
      const raw = localStorage.getItem(ATTEMPTS_KEY);
      return raw ? JSON.parse(raw) : [];
    } catch (e) {
      return [];
    }
  }

  function saveLocalAttempt(attempt) {
    const list = loadLocalAttempts();
    list.unshift(attempt);
    localStorage.setItem(ATTEMPTS_KEY, JSON.stringify(list));
  }

  // ─── SUPABASE / HYBRID API ────────────────────────────────
  const KuisDB = {
    async fetchQuizzes() {
      let quizzes = [];
      if (typeof SUPABASE_URL !== 'undefined' && typeof SUPABASE_ANON_KEY !== 'undefined') {
        try {
          const res = await fetch(`${SUPABASE_URL}/rest/v1/quizzes?select=*,quiz_questions(*)&order=created_at.desc`, {
            headers: SUPABASE_HEADERS
          });
          if (res.ok) {
            const data = await res.json();
            if (Array.isArray(data) && data.length > 0) {
              quizzes = data.map(q => ({
                ...q,
                questions: (q.quiz_questions || []).sort((a,b) => (a.order_seq || 0) - (b.order_seq || 0))
              }));
            }
          }
        } catch (e) {
          console.warn('[KuisDB] Supabase fetch failed, falling back to local:', e);
        }
      }
      if (!quizzes || quizzes.length === 0) {
        quizzes = loadLocalQuizzes();
      }
      return quizzes;
    },

    async saveQuiz(quizData) {
      // Check if updating or creating
      const isNew = !quizData.id || String(quizData.id).startsWith('quiz-');
      const quizId = isNew ? (crypto.randomUUID ? crypto.randomUUID() : 'quiz-' + Date.now()) : quizData.id;
      
      const payload = {
        id: quizId,
        title: quizData.title,
        description: quizData.description || '',
        category: quizData.category || 'Umum',
        pin_code: quizData.pin_code || String(Math.floor(100000 + Math.random() * 900000)),
        is_active: quizData.is_active !== false,
        time_per_question_sec: parseInt(quizData.time_per_question_sec) || 20,
        created_by: quizData.created_by || 'Admin',
        updated_at: new Date().toISOString()
      };

      // Try Supabase first
      let supabaseSuccess = false;
      if (typeof SUPABASE_URL !== 'undefined' && typeof SUPABASE_ANON_KEY !== 'undefined') {
        try {
          const res = await fetch(`${SUPABASE_URL}/rest/v1/quizzes`, {
            method: 'POST',
            headers: {
              ...SUPABASE_HEADERS,
              'Prefer': 'resolution=merge-duplicates'
            },
            body: JSON.stringify(payload)
          });
          if (res.ok) {
            // Delete old questions and insert new ones
            await fetch(`${SUPABASE_URL}/rest/v1/quiz_questions?quiz_id=eq.${quizId}`, {
              method: 'DELETE',
              headers: SUPABASE_HEADERS
            });

            if (Array.isArray(quizData.questions) && quizData.questions.length > 0) {
              const questionsPayload = quizData.questions.map((q, idx) => ({
                id: q.id && !String(q.id).startsWith('q') ? q.id : (crypto.randomUUID ? crypto.randomUUID() : 'q-' + Date.now() + '-' + idx),
                quiz_id: quizId,
                order_seq: idx + 1,
                question_text: q.question_text,
                option_a: q.option_a,
                option_b: q.option_b,
                option_c: q.option_c,
                option_d: q.option_d,
                correct_option: q.correct_option,
                points: parseInt(q.points) || 100
              }));

              await fetch(`${SUPABASE_URL}/rest/v1/quiz_questions`, {
                method: 'POST',
                headers: SUPABASE_HEADERS,
                body: JSON.stringify(questionsPayload)
              });
            }
            supabaseSuccess = true;
          }
        } catch (e) {
          console.warn('[KuisDB] Supabase save error:', e);
        }
      }

      // Always update local storage too
      const localList = loadLocalQuizzes();
      const updatedQuestions = (quizData.questions || []).map((q, idx) => ({
        id: q.id || 'q-' + Date.now() + '-' + idx,
        order_seq: idx + 1,
        question_text: q.question_text,
        option_a: q.option_a,
        option_b: q.option_b,
        option_c: q.option_c,
        option_d: q.option_d,
        correct_option: q.correct_option,
        points: parseInt(q.points) || 100
      }));

      const fullQuizObj = { ...payload, questions: updatedQuestions };

      const existingIdx = localList.findIndex(x => String(x.id) === String(quizId));
      if (existingIdx >= 0) {
        localList[existingIdx] = fullQuizObj;
      } else {
        localList.unshift(fullQuizObj);
      }
      saveLocalQuizzes(localList);

      return fullQuizObj;
    },

    async deleteQuiz(quizId) {
      if (typeof SUPABASE_URL !== 'undefined' && typeof SUPABASE_ANON_KEY !== 'undefined') {
        try {
          await fetch(`${SUPABASE_URL}/rest/v1/quizzes?id=eq.${quizId}`, {
            method: 'DELETE',
            headers: SUPABASE_HEADERS
          });
        } catch (e) {}
      }
      const localList = loadLocalQuizzes().filter(x => String(x.id) !== String(quizId));
      saveLocalQuizzes(localList);
      return true;
    },

    async saveAttempt(attemptObj) {
      const payload = {
        id: crypto.randomUUID ? crypto.randomUUID() : 'att-' + Date.now(),
        quiz_id: attemptObj.quiz_id,
        user_nip: attemptObj.user_nip,
        user_name: attemptObj.user_name,
        total_score: attemptObj.total_score,
        correct_count: attemptObj.correct_count,
        wrong_count: attemptObj.wrong_count,
        max_streak: attemptObj.max_streak || 0,
        completed_at: new Date().toISOString()
      };

      if (typeof SUPABASE_URL !== 'undefined' && typeof SUPABASE_ANON_KEY !== 'undefined') {
        try {
          await fetch(`${SUPABASE_URL}/rest/v1/quiz_attempts`, {
            method: 'POST',
            headers: SUPABASE_HEADERS,
            body: JSON.stringify(payload)
          });
        } catch (e) {}
      }

      saveLocalAttempt(payload);
      return payload;
    },

    async fetchAttempts(quizId) {
      let attempts = [];
      if (typeof SUPABASE_URL !== 'undefined' && typeof SUPABASE_ANON_KEY !== 'undefined') {
        try {
          const res = await fetch(`${SUPABASE_URL}/rest/v1/quiz_attempts?quiz_id=eq.${quizId}&order=total_score.desc,completed_at.asc`, {
            headers: SUPABASE_HEADERS
          });
          if (res.ok) {
            attempts = await res.json();
          }
        } catch (e) {}
      }
      if (!attempts || attempts.length === 0) {
        attempts = loadLocalAttempts().filter(x => String(x.quiz_id) === String(quizId));
        attempts.sort((a,b) => (b.total_score || 0) - (a.total_score || 0));
      }
      return attempts;
    }
  };

  // ─── SCORING ENGINE ───────────────────────────────────────
  function calculateScore(basePoints, timeRemainingSec, totalTimeSec, currentStreak) {
    if (timeRemainingSec <= 0) return 0;
    const speedRatio = Math.max(0, timeRemainingSec / totalTimeSec);
    // Base formula: 50% for correct answer + up to 50% bonus for speed
    let points = Math.round(basePoints * (0.5 + (0.5 * speedRatio)));
    
    // Streak multiplier
    let streakBonus = 1;
    if (currentStreak >= 3) streakBonus = 1.5;
    else if (currentStreak >= 2) streakBonus = 1.25;

    return Math.round(points * streakBonus);
  }

  // EXPOSE TO GLOBAL SCOPE
  global.KuisAudio = KuisAudio;
  global.KuisDB = KuisDB;
  global.calculateScore = calculateScore;

})(typeof window !== 'undefined' ? window : this);
