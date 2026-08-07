-- ═══════════════════════════════════════════════════════════════════════════
-- SUPABASE MIGRATION: Fitur Arena Kuis Interaktif (Kahoot / Wayground Style)
-- ═══════════════════════════════════════════════════════════════════════════

-- 1. TABEL QUIZZES (Daftar Kuis)
CREATE TABLE IF NOT EXISTS public.quizzes (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    title VARCHAR(255) NOT NULL,
    description TEXT,
    category VARCHAR(100) DEFAULT 'Umum',
    pin_code VARCHAR(10),
    is_active BOOLEAN DEFAULT true,
    time_per_question_sec INT DEFAULT 20,
    created_by VARCHAR(100),
    created_at TIMESTAMPTZ DEFAULT NOW(),
    updated_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. TABEL QUIZ_QUESTIONS (Pertanyaan Kuis)
CREATE TABLE IF NOT EXISTS public.quiz_questions (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    quiz_id UUID REFERENCES public.quizzes(id) ON DELETE CASCADE,
    order_seq INT DEFAULT 1,
    question_text TEXT NOT NULL,
    option_a TEXT NOT NULL,
    option_b TEXT NOT NULL,
    option_c TEXT NOT NULL,
    option_d TEXT NOT NULL,
    correct_option VARCHAR(1) NOT NULL CHECK (correct_option IN ('A', 'B', 'C', 'D')),
    points INT DEFAULT 100,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 3. TABEL QUIZ_ATTEMPTS (Sesi Pengerjaan Peserta)
CREATE TABLE IF NOT EXISTS public.quiz_attempts (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    quiz_id UUID REFERENCES public.quizzes(id) ON DELETE CASCADE,
    user_nip VARCHAR(100) NOT NULL,
    user_name VARCHAR(255) NOT NULL,
    total_score INT DEFAULT 0,
    correct_count INT DEFAULT 0,
    wrong_count INT DEFAULT 0,
    max_streak INT DEFAULT 0,
    completed_at TIMESTAMPTZ DEFAULT NOW()
);

-- 4. TABEL QUIZ_RESPONSES (Jawaban Per Soal Peserta)
CREATE TABLE IF NOT EXISTS public.quiz_responses (
    id UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    attempt_id UUID REFERENCES public.quiz_attempts(id) ON DELETE CASCADE,
    question_id UUID REFERENCES public.quiz_questions(id) ON DELETE CASCADE,
    selected_option VARCHAR(1) CHECK (selected_option IN ('A', 'B', 'C', 'D')),
    is_correct BOOLEAN DEFAULT false,
    response_time_sec NUMERIC(5,2) DEFAULT 0,
    score_earned INT DEFAULT 0,
    created_at TIMESTAMPTZ DEFAULT NOW()
);

-- RLS POLICIES (Bypass anon & authenticated for ease of use)
ALTER TABLE public.quizzes ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.quiz_questions ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.quiz_attempts ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.quiz_responses ENABLE ROW LEVEL SECURITY;

DROP POLICY IF EXISTS "Public select quizzes" ON public.quizzes;
CREATE POLICY "Public select quizzes" ON public.quizzes FOR SELECT USING (true);
DROP POLICY IF EXISTS "Public insert quizzes" ON public.quizzes;
CREATE POLICY "Public insert quizzes" ON public.quizzes FOR INSERT WITH CHECK (true);
DROP POLICY IF EXISTS "Public update quizzes" ON public.quizzes;
CREATE POLICY "Public update quizzes" ON public.quizzes FOR UPDATE USING (true);
DROP POLICY IF EXISTS "Public delete quizzes" ON public.quizzes;
CREATE POLICY "Public delete quizzes" ON public.quizzes FOR DELETE USING (true);

DROP POLICY IF EXISTS "Public select questions" ON public.quiz_questions;
CREATE POLICY "Public select questions" ON public.quiz_questions FOR SELECT USING (true);
DROP POLICY IF EXISTS "Public insert questions" ON public.quiz_questions;
CREATE POLICY "Public insert questions" ON public.quiz_questions FOR INSERT WITH CHECK (true);
DROP POLICY IF EXISTS "Public update questions" ON public.quiz_questions;
CREATE POLICY "Public update questions" ON public.quiz_questions FOR UPDATE USING (true);
DROP POLICY IF EXISTS "Public delete questions" ON public.quiz_questions;
CREATE POLICY "Public delete questions" ON public.quiz_questions FOR DELETE USING (true);

DROP POLICY IF EXISTS "Public select attempts" ON public.quiz_attempts;
CREATE POLICY "Public select attempts" ON public.quiz_attempts FOR SELECT USING (true);
DROP POLICY IF EXISTS "Public insert attempts" ON public.quiz_attempts;
CREATE POLICY "Public insert attempts" ON public.quiz_attempts FOR INSERT WITH CHECK (true);

DROP POLICY IF EXISTS "Public select responses" ON public.quiz_responses;
CREATE POLICY "Public select responses" ON public.quiz_responses FOR SELECT USING (true);
DROP POLICY IF EXISTS "Public insert responses" ON public.quiz_responses;
CREATE POLICY "Public insert responses" ON public.quiz_responses FOR INSERT WITH CHECK (true);

-- SEED DATA CONTOH KUIS
INSERT INTO public.quizzes (id, title, description, category, pin_code, is_active, time_per_question_sec, created_by)
VALUES 
(
  'e1111111-1111-1111-1111-111111111111',
  'Kuis Pengetahuan Umum Portal 9201',
  'Uji wawasan Anda seputar fitur, regulasi kepegawaian, dan budaya kerja Portal 9201 secara interaktif!',
  'Kepegawaian',
  '920101',
  true,
  20,
  'Admin'
) ON CONFLICT (id) DO NOTHING;

INSERT INTO public.quiz_questions (quiz_id, order_seq, question_text, option_a, option_b, option_c, option_d, correct_option, points)
VALUES
(
  'e1111111-1111-1111-1111-111111111111',
  1,
  'Apa kepanjangan dari fitur EoTQ pada Portal 9201?',
  'Employee of The Quarter',
  'Evaluation of The Quality',
  'Executive of The Quarter',
  'Employee of The Quantity',
  'A',
  100
),
(
  'e1111111-1111-1111-1111-111111111111',
  2,
  'Dokumen apa yang biasanya diminta oleh pegawai sebelum melakukan perjalanan dinas?',
  'Kamus POK',
  'Surat Tugas',
  'Nota Beban Work',
  'Sertifikat Pelatihan',
  'B',
  100
),
(
  'e1111111-1111-1111-1111-111111111111',
  3,
  'Warna khas aksen utama yang menghiasi tema Portal 9201 adalah kombinasi dari...',
  'Merah & Hitam',
  'Hijau & Perak',
  'Navy (#0d2340) & Gold (#c8a84b)',
  'Kuning & Ungu',
  'C',
  100
) ON CONFLICT DO NOTHING;
