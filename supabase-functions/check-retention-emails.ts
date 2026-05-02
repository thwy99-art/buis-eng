import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'
import { createClient } from 'https://esm.sh/@supabase/supabase-js@2'

const SUPABASE_URL = Deno.env.get('SUPABASE_URL')!
const SUPABASE_SERVICE_ROLE_KEY = Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!
const SEND_EMAIL_URL = `${SUPABASE_URL}/functions/v1/quick-endpoint`

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const { type } = await req.json()

  const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY)
  const now = new Date()

  const results: string[] = []

  if (type === 'daily') {
    // 스트릭 위기 알림: 마지막 학습 24~48시간 전
    const h24ago = new Date(now.getTime() - 24 * 60 * 60 * 1000).toISOString()
    const h48ago = new Date(now.getTime() - 48 * 60 * 60 * 1000).toISOString()

    const { data: reminderUsers } = await supabase
      .from('profiles')
      .select('email, lesson_count')
      .eq('email_notifications', true)
      .not('last_lesson_at', 'is', null)
      .gte('last_lesson_at', h48ago)
      .lt('last_lesson_at', h24ago)

    for (const user of reminderUsers ?? []) {
      await sendEmail('streak-reminder', user.email, { lessonCount: user.lesson_count })
      results.push(`streak-reminder → ${user.email}`)
    }

    // 스트릭 부활 알림: 마지막 학습 72~96시간 전
    const h72ago = new Date(now.getTime() - 72 * 60 * 60 * 1000).toISOString()
    const h96ago = new Date(now.getTime() - 96 * 60 * 60 * 1000).toISOString()

    const { data: revivalUsers } = await supabase
      .from('profiles')
      .select('email, lesson_count')
      .eq('email_notifications', true)
      .not('last_lesson_at', 'is', null)
      .gte('last_lesson_at', h96ago)
      .lt('last_lesson_at', h72ago)

    for (const user of revivalUsers ?? []) {
      await sendEmail('streak-revival', user.email, { lessonCount: user.lesson_count })
      results.push(`streak-revival → ${user.email}`)
    }

  } else if (type === 'weekly') {
    // 주간 리포트: 학습 기록이 1개 이상 + 마지막 학습 30일 이내
    const days30ago = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000).toISOString()

    const { data: weeklyUsers } = await supabase
      .from('profiles')
      .select('email, lesson_count')
      .eq('email_notifications', true)
      .gt('lesson_count', 0)
      .gte('last_lesson_at', days30ago)

    for (const user of weeklyUsers ?? []) {
      await sendEmail('weekly-report', user.email, { lessonCount: user.lesson_count })
      results.push(`weekly-report → ${user.email}`)
    }
  }

  return new Response(JSON.stringify({ sent: results.length, results }), {
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    status: 200,
  })
})

async function sendEmail(type: string, email: string, data: Record<string, unknown> = {}) {
  await fetch(SEND_EMAIL_URL, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${SUPABASE_SERVICE_ROLE_KEY}`,
    },
    body: JSON.stringify({ type, email, ...data }),
  })
}
