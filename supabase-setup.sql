-- ① lessons 테이블 생성
create table if not exists public.lessons (
  id          bigserial primary key,
  job         text not null check (job in ('designer','marketer','developer','pm')),
  difficulty  text not null check (difficulty in ('기초','중급','고급')),
  situation   text not null,
  correct     text not null,
  correct_ko  text,
  wrong1      text not null,
  wrong2      text not null,
  reason      text not null,
  is_premium  boolean not null default false,
  created_at  timestamptz default now()
);

-- ② 인덱스 (직무별 fetch 빠르게)
create index if not exists lessons_job_idx on public.lessons (job);
create index if not exists lessons_premium_idx on public.lessons (job, is_premium);

-- ③ RLS 활성화
alter table public.lessons enable row level security;

-- ④ 무료 문항: 누구나 읽기 가능 (비로그인 포함)
create policy "free lessons are public"
  on public.lessons for select
  using (is_premium = false);

-- ⑤ 유료 문항: 로그인한 유저만
create policy "premium lessons require auth"
  on public.lessons for select
  using (
    is_premium = false
    or auth.role() = 'authenticated'
  );
