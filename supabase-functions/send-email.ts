import { serve } from 'https://deno.land/std@0.168.0/http/server.ts'

const RESEND_API_KEY = Deno.env.get('RESEND_API_KEY')!
const FROM = '3분 업무 영어 <onboarding@resend.dev>' // 도메인 연결 전 테스트용
const APP_URL = 'https://thwy99-art.github.io/buis-eng'

const corsHeaders = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type',
}

serve(async (req) => {
  if (req.method === 'OPTIONS') {
    return new Response('ok', { headers: corsHeaders })
  }

  const body = await req.json()
  const { type, email, lessonCount = 0 } = body

  let subject = ''
  let html = ''

  // ── 공통 스타일 ──
  const wrap = (content: string) => `
    <div style="font-family:'Apple SD Gothic Neo',Pretendard,sans-serif;max-width:480px;margin:0 auto;padding:40px 24px;color:#111118;background:#fff">
      ${content}
      <hr style="border:none;border-top:1px solid #e8e8f0;margin:32px 0">
      <p style="color:#9999aa;font-size:12px;line-height:1.6;margin:0">
        3분 업무 영어 · 이메일 알림을 끄려면 앱 내 마이페이지 → 이메일 알림에서 설정하세요.
      </p>
    </div>
  `
  const cta = (text: string) =>
    `<a href="${APP_URL}" style="display:inline-block;background:#5b4fff;color:white;padding:14px 28px;border-radius:12px;text-decoration:none;font-weight:700;font-size:15px">${text}</a>`

  // ── 이메일 타입별 콘텐츠 ──

  if (type === 'welcome') {
    subject = '3분 업무 영어에 오신 걸 환영해요! 🎉'
    html = wrap(`
      <h1 style="font-size:24px;font-weight:800;margin-bottom:8px">환영해요! 🎉</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:24px">
        3분 업무 영어에 가입해주셔서 감사해요.<br>
        매일 3분, 직무별 맞춤 표현으로 자연스러운 업무 영어를 만들어드릴게요.
      </p>
      ${cta('첫 레슨 시작하기 →')}
    `)

  } else if (type === 'withdrawal') {
    subject = '3분 업무 영어 탈퇴가 완료됐어요'
    html = wrap(`
      <h1 style="font-size:24px;font-weight:800;margin-bottom:8px">탈퇴가 완료됐어요</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7">
        그동안 3분 업무 영어를 이용해주셔서 감사했어요.<br>
        모든 계정 정보와 학습 기록이 삭제됐어요.
      </p>
      <p style="color:#666680;font-size:13px;margin-top:16px">언제든 다시 돌아오시면 반갑게 맞이할게요.</p>
    `)

  } else if (type === 'streak-reminder') {
    subject = '오늘 영어 연습, 아직 안 하셨어요 🔥'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">📌 학습 알림</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">아직 오늘 연습을<br>못 하셨어요</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:8px">
        딱 3분이면 오늘 학습을 완료할 수 있어요.<br>
        이동 중이나 점심시간에 짧게 하셔도 충분합니다.
      </p>
      <p style="color:#444;font-size:14px;font-weight:600;margin-bottom:24px">누적 학습 ${lessonCount}회 — 지금까지 쌓은 거 아깝잖아요 💪</p>
      ${cta('3분 연습하러 가기 →')}
    `)

  } else if (type === 'streak-revival') {
    subject = '3분 업무 영어가 기다리고 있었어요'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">💌 다시 시작해요</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">며칠 동안<br>연습이 없었네요</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:8px">
        바쁘셨죠? 충분히 이해해요.<br>
        그래도 영어 실력은 조금씩 매일이 훨씬 효과적이거든요.
      </p>
      <p style="color:#444;font-size:14px;font-weight:600;margin-bottom:24px">오늘 딱 1문제만 풀어도 다시 시작하신 거예요.</p>
      ${cta('오늘의 레슨 시작하기 →')}
      <p style="color:#9999aa;font-size:13px;margin-top:20px">며칠 쉬었어도 괜찮아요. 오늘이 새로운 Day 1입니다.</p>
    `)

  } else if (type === 'weekly-report') {
    subject = '이번 주도 영어 연습, 잊지 않으셨죠? 📊'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">📊 주간 리포트</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">이번 주 목표를<br>세워볼까요?</h1>
      <div style="background:#f5f5ff;border-radius:12px;padding:20px;margin-bottom:24px">
        <p style="margin:0;color:#666680;font-size:13px">누적 학습</p>
        <p style="margin:4px 0 0;font-size:28px;font-weight:800;color:#5b4fff">${lessonCount}회</p>
        <p style="margin:8px 0 0;color:#444;font-size:14px">월~금, 하루 3분씩이면<br>이번 주 5개 레슨 달성 가능합니다.</p>
      </div>
      ${cta('이번 주 첫 레슨 시작하기 →')}
    `)

  } else if (type === 'milestone-10') {
    subject = '와, 10개 완료! 벌써 상위 30%예요 🎉'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">🏆 마일스톤 달성</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">10개 레슨 완료!</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:8px">
        꾸준히 하시는 분이 드물어요.<br>
        이 정도면 정말 진심이신 거예요.
      </p>
      <p style="color:#444;font-size:14px;font-weight:600;margin-bottom:24px">오늘 배운 표현, 내일 회의에서 한 번 써보세요 ✦</p>
      ${cta('오늘 레슨 바로가기 →')}
    `)

  } else if (type === 'milestone-30') {
    subject = '한 달의 기적 — 30레슨 달성! 🔥'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">🏆 마일스톤 달성</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">30개 레슨 완료!</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:8px">
        30개면 이미 직장 영어의 핵심 표현들을 다 익히셨어요.<br>
        원어민 동료와 미팅해도 당황하지 않을 수준이에요.
      </p>
      <p style="color:#444;font-size:14px;font-weight:600;margin-bottom:24px">다음 목표는 100개 — 계속 해봐요 💪</p>
      ${cta('계속 학습하기 →')}
    `)

  } else if (type === 'milestone-100') {
    subject = '100레슨 완주! 당신은 진짜입니다 🏅'
    html = wrap(`
      <p style="font-size:13px;font-weight:600;color:#5b4fff;letter-spacing:1px;margin-bottom:12px">🏅 전설의 마일스톤</p>
      <h1 style="font-size:22px;font-weight:800;margin-bottom:8px;line-height:1.4">100개 레슨 완료!</h1>
      <p style="color:#666680;font-size:15px;line-height:1.7;margin-bottom:8px">
        100개를 완주한 분은 전체 이용자 중 상위 1%예요.<br>
        매일 3분을 지켜온 당신의 노력이 정말 대단합니다.
      </p>
      <p style="color:#444;font-size:14px;font-weight:600;margin-bottom:24px">이 성취, 팀원들에게 자랑해보세요 ✦</p>
      ${cta('다음 레슨 이어가기 →')}
    `)
  }

  if (!subject) {
    return new Response(JSON.stringify({ error: 'unknown type' }), {
      headers: { ...corsHeaders, 'Content-Type': 'application/json' },
      status: 400,
    })
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
  if (!res.ok) console.error('Resend error:', JSON.stringify(data))
  return new Response(JSON.stringify(data), {
    headers: { ...corsHeaders, 'Content-Type': 'application/json' },
    status: res.ok ? 200 : 400,
  })
})
