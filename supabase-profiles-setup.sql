-- ══ profiles 테이블 ══
-- 이메일 알림 설정, 마지막 학습일, 누적 학습 수를 서버에 저장

create table if not exists public.profiles (
  id                    uuid references auth.users(id) on delete cascade primary key,
  email                 text,
  email_notifications   boolean not null default true,
  last_lesson_at        timestamptz,
  lesson_count          integer not null default 0,
  created_at            timestamptz default now(),
  updated_at            timestamptz default now()
);

-- RLS 활성화
alter table public.profiles enable row level security;

-- 본인 프로필만 읽기/수정 가능
create policy "users can view own profile"
  on public.profiles for select
  using (auth.uid() = id);

create policy "users can update own profile"
  on public.profiles for update
  using (auth.uid() = id);

create policy "users can insert own profile"
  on public.profiles for insert
  with check (auth.uid() = id);

-- 서비스 롤 전용: 예약 이메일 발송 시 전체 조회
create policy "service role can read all profiles"
  on public.profiles for select
  using (auth.role() = 'service_role');

-- 회원 가입 시 자동 프로필 생성 함수
create or replace function public.handle_new_user()
returns trigger as $$
begin
  insert into public.profiles (id, email)
  values (new.id, new.email)
  on conflict (id) do nothing;
  return new;
end;
$$ language plpgsql security definer;

-- 가입 트리거
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();


-- ══ pg_cron 예약 작업 설정 ══
-- Supabase 대시보드 > Database > Extensions 에서 pg_cron, pg_net 활성화 필요
-- 아래 YOUR_SERVICE_ROLE_KEY 를 실제 서비스 롤 키로 교체 후 실행

/*
-- 매일 15:00 UTC (자정 KST) — 스트릭 알림 체크
select cron.schedule(
  'daily-retention-emails',
  '0 15 * * *',
  $$
  select net.http_post(
    url    := 'https://yitssseegmpjjjsusowp.supabase.co/functions/v1/check-retention-emails',
    headers:= '{"Authorization":"Bearer YOUR_SERVICE_ROLE_KEY","Content-Type":"application/json"}'::jsonb,
    body   := '{"type":"daily"}'::jsonb
  );
  $$
);

-- 매주 일요일 23:00 UTC (월요일 8:00 KST) — 주간 리포트
select cron.schedule(
  'weekly-retention-email',
  '0 23 * * 0',
  $$
  select net.http_post(
    url    := 'https://yitssseegmpjjjsusowp.supabase.co/functions/v1/check-retention-emails',
    headers:= '{"Authorization":"Bearer YOUR_SERVICE_ROLE_KEY","Content-Type":"application/json"}'::jsonb,
    body   := '{"type":"weekly"}'::jsonb
  );
  $$
);
*/
