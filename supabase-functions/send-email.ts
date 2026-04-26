import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'

const RESEND_API_KEY = Deno.env.get('RESEND_API_KEY')!
const FROM = '3분 업무 영어 <noreply@yourdomain.com>' // 도메인 설정 후 변경

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const { type, email } = await req.json()

  let subject = ''
  let html = ''

  if (type === 'welcome') {
    subject = '3분 업무 영어에 오신 걸 환영해요! 🎉'
    html = `
      <div style="font-family:'Apple SD Gothic Neo',sans-serif;max-width:480px;margin:0 auto;padding:40px 24px;color:#111118">
        <h1 style="font-size:24px;font-weight:800;margin-bottom:8px">환영해요! 🎉</h1>
        <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:24px">
          3분 업무 영어에 가입해주셔서 감사해요.<br>
          매일 3분, 직무별 맞춤 표현으로 자연스러운 업무 영어를 만들어드릴게요.
        </p>
        <a href="http://localhost:8080" style="display:inline-block;background:#5b4fff;color:white;padding:14px 28px;border-radius:12px;text-decoration:none;font-weight:700;font-size:15px">
          학습 시작하기 →
        </a>
      </div>
    `
  } else if (type === 'withdrawal') {
    subject = '3분 업무 영어 탈퇴가 완료됐어요'
    html = `
      <div style="font-family:'Apple SD Gothic Neo',sans-serif;max-width:480px;margin:0 auto;padding:40px 24px;color:#111118">
        <h1 style="font-size:24px;font-weight:800;margin-bottom:8px">탈퇴가 완료됐어요</h1>
        <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:24px">
          그동안 3분 업무 영어를 이용해주셔서 감사했어요.<br>
          모든 계정 정보와 학습 기록이 삭제됐어요.
        </p>
        <p style="color:#666680;font-size:13px">
          언제든 다시 돌아오시면 반갑게 맞이할게요.
        </p>
      </div>
    `
  }

  const res = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${RESEND_API_KEY}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ from: FROM, to: email, subject, html }),
  })

  const data = await res.json()
  return new Response(JSON.stringify(data), {
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    status: res.ok ? 200 : 400,
  })
})

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}
