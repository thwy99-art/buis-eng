import xlsxwriter

# (난이도, 상황(한국어 문장), 어색한표현→wrong1, 정답, wrong2, 정답이유(문장형))
data = [
    # ── 기초 20 ──
    ("기초", "성과 보고 미팅에서 지난달 캠페인 결과를 팀에 공유하고 싶어요.",
     "The campaign did good results last month.",
     "The campaign delivered strong results last month.",
     "The campaign made good results last month.",
     "'do results'는 비문이에요. 'deliver'는 성과나 결과를 달성하다는 의미의 비즈니스 표현이에요."),

    ("기초", "성과 보고에서 전환율 변화를 수치로 설명하고 싶어요.",
     "Our conversion rate is going up 15%.",
     "Our conversion rate increased by 15%.",
     "Our conversion rate is going up with 15%.",
     "'going up 15%'는 어색한 표현이에요. 'increased by + 수치'가 정확한 비즈니스 표현이에요."),

    ("기초", "KPI 보고에서 목표 달성에 미치지 못했다고 솔직하게 말하고 싶어요.",
     "We didn't reach our goal of 10,000 signups.",
     "We fell short of our target of 10,000 signups.",
     "We didn't arrive our goal of 10,000 signups.",
     "'fell short of'는 목표에 미달했다는 가장 자연스러운 비즈니스 표현이에요."),

    ("기초", "동료에게 이메일로 보고서를 금요일까지 보내달라고 요청하고 싶어요.",
     "Can you send me the report until Friday?",
     "Can you send me the report by Friday?",
     "Can you send me the report before to Friday?",
     "마감 기한에는 'by'를 써요. 'until'은 그 시점까지 계속 지속되는 상황에 쓰는 표현이에요."),

    ("기초", "동료에게 캠페인 계획에 대해 이야기하자고 미팅을 요청하고 싶어요.",
     "I want to discuss about the campaign plan.",
     "I want to discuss the campaign plan.",
     "I want to discuss on the campaign plan.",
     "'discuss'는 타동사라서 전치사 없이 바로 목적어를 써요. 'discuss about'은 문법 오류예요."),

    ("기초", "지난 분기 광고비 지출 내역을 팀에 보고하고 싶어요.",
     "We spend a lot of money for ads last quarter.",
     "We spent a significant budget on ads last quarter.",
     "We spending a lot of money for ads last quarter.",
     "'spent'이 올바른 과거형이에요. 'significant budget on ads'가 비즈니스 맥락에서 더 자연스러운 표현이에요."),

    ("기초", "금요일 오후에 미팅이 가능한지 동료에게 확인하고 싶어요.",
     "Are you free on this Friday afternoon?",
     "Are you available this Friday afternoon?",
     "Are you having time on this Friday afternoon?",
     "비즈니스에서는 'available'이 'free'보다 공식적인 표현이에요. 'on this Friday'에서 'on'은 생략하는 게 자연스러워요."),

    ("기초", "미팅에서 광고 예산을 늘리자고 제안하고 싶어요.",
     "I think we should to increase the budget.",
     "I think we should increase the budget.",
     "I think we should increasing the budget.",
     "'should' 뒤에는 동사원형을 써요. 'to'는 필요 없어요."),

    ("기초", "프로세스를 잘 모르는 외부 담당자에게 설명을 요청하고 싶어요.",
     "Could you explain me the process?",
     "Could you explain the process to me?",
     "Could you explain for me the process?",
     "'explain'은 'explain something to someone' 구조로 써요. 'explain me'는 문법 오류예요."),

    ("기초", "동료에게 광고 카피에 대한 피드백을 부탁하고 싶어요.",
     "Please give me your opinion about this copy.",
     "I'd love to get your feedback on this copy.",
     "Please tell me your opinion about this copy.",
     "'feedback on'이 카피 리뷰 요청에 자연스러운 표현이에요. 'I'd love to'를 쓰면 공손한 인상을 줘요."),

    ("기초", "미팅에서 동료의 의견에 동의를 표현하고 싶어요.",
     "I'm agree with your point.",
     "I agree with your point.",
     "I'm agreeing with your point.",
     "'agree'는 동사라서 'be동사 + agree' 형태는 문법 오류예요. 'I agree'가 올바른 표현이에요."),

    ("기초", "미팅에서 상대방의 제안 방향에 부드럽게 반대하고 싶어요.",
     "I don't agree that idea.",
     "I don't think that's the right approach.",
     "I disagree about that idea.",
     "직접 반대보다 'I don't think that's the right approach'처럼 표현하면 비즈니스적으로 더 자연스러워요."),

    ("기초", "다음 주로 미팅 일정을 변경하고 싶어요.",
     "Can we push the meeting to next week?",
     "Can we reschedule the meeting to next week?",
     "Can we move back the meeting to next week?",
     "'reschedule'이 미팅 일정 변경의 표준 표현이에요. 전치사는 'to'를 써요."),

    ("기초", "작업이 완료되면 알려달라고 동료에게 요청하고 싶어요.",
     "Please confirm to me when it's ready.",
     "Please let me know when it's ready.",
     "Please tell to me when it's ready.",
     "'let me know'가 확인 요청에서 가장 자연스럽고 흔히 쓰는 표현이에요."),

    ("기초", "협조해준 파트너에게 감사 인사를 전하고 싶어요.",
     "Thank you for your corporation.",
     "Thank you for your cooperation.",
     "Thank you for your collaborate.",
     "'cooperation'은 협조, 'corporation'은 기업을 뜻해요. 두 단어를 혼동하지 않도록 주의하세요."),

    ("기초", "동료에게 파일을 공유하겠다고 말하고 싶어요.",
     "I will share you the file.",
     "I'll share the file with you.",
     "I will share to you the file.",
     "'share'는 'share A with B' 구조로 써요. 'share you the file'은 문법 오류예요."),

    ("기초", "광고 CTR 수치를 보고서에서 정확하게 표현하고 싶어요.",
     "The CTR of the ad is 3.2 percents.",
     "The CTR of the ad is 3.2%.",
     "The CTR of the ad are 3.2%.",
     "퍼센트는 '%' 기호를 쓰거나 단수형 'percent'를 써요. 'percents'는 비문이에요."),

    ("기초", "발표를 마무리하며 핵심 내용을 정리하고 싶어요.",
     "For summarize, we achieved most of our goals.",
     "In summary, we met most of our targets.",
     "To conclude, we achieved most of our goals.",
     "'For summarize'는 문법 오류예요. 전치사 뒤에는 동명사를 써야 해서 'For summarizing'이 맞고, 발표 마무리에는 'In summary'가 가장 자연스러운 표현이에요."),

    ("기초", "팀 미팅에서 새로운 방향을 부드럽게 제안하고 싶어요.",
     "Why don't we try new approach?",
     "Why don't we try a new approach?",
     "Why don't we trying new approach?",
     "관사 'a'가 필요해요. 어순은 'Why don't we + 동사원형'이에요."),

    ("기초", "이메일에서 마감일을 정확하게 전달하고 싶어요.",
     "The deadline is at April 30th.",
     "The deadline is April 30th.",
     "The deadline will be on April 30th.",
     "날짜 앞에는 'at'이 아닌 'on'을 써요. 비즈니스 이메일에서는 'The deadline is April 30th'처럼 전치사를 생략하는 게 가장 간결하고 자연스러워요."),

    # ── 중급 50 ──
    ("중급", "A/B 테스트 결과를 팀에 비교해서 보고하고 싶어요.",
     "The A version performed better than B version.",
     "Version A outperformed Version B by a significant margin.",
     "Version A was better than Version B.",
     "'outperformed by a significant margin'이 A/B 테스트 결과 보고에 더 정확하고 임팩트 있는 표현이에요."),

    ("중급", "데이터 발표에서 참여율 상승 트렌드를 설명하고 싶어요.",
     "Our data shows an increasing trend.",
     "The data indicates an upward trend in engagement.",
     "Our data is showing an increasing trend.",
     "'indicates'가 데이터 분석에 정확한 동사예요. 'upward trend in engagement'처럼 구체적으로 표현하는 게 좋아요."),

    ("중급", "캠페인이 예상보다 저조한 성과를 냈다고 보고하고 싶어요.",
     "The campaign didn't work as we planned.",
     "The campaign underperformed against our projections.",
     "The campaign was not successful as planned.",
     "'underperformed against projections'가 목표 대비 결과를 전문적이고 객관적으로 설명하는 표현이에요."),

    ("중급", "임원에게 캠페인의 투자 수익률을 설명하고 싶어요.",
     "This campaign has a good return.",
     "This campaign generated a 3x return on investment.",
     "This campaign shows good ROI.",
     "구체적인 수치(3x)와 'generated'를 함께 쓰면 임원 보고에서 훨씬 설득력 있어요."),

    ("중급", "소셜 미디어 채널에 더 집중하자고 전략적으로 제안하고 싶어요.",
     "I suggest to focus more on social media.",
     "I'd recommend shifting our focus toward social media channels.",
     "I propose we focus more on social media.",
     "'I'd recommend'가 제안할 때 공손하고 전문적인 표현이에요. 'shifting focus toward'는 전략적 변화를 잘 표현해요."),

    ("중급", "다음 분기 예산 증액을 경영진에게 요청하고 싶어요.",
     "We need more budget for next quarter.",
     "We're requesting an additional budget allocation for Q3 to scale high-performing campaigns.",
     "We need extra budget in the next quarter.",
     "구체적인 이유(scale high-performing campaigns)를 포함한 예산 요청이 승인 가능성을 높여요."),

    ("중급", "외부 파트너에게 협업 가능성을 타진하고 싶어요.",
     "We want to work together with your company.",
     "We'd love to explore a partnership opportunity with your team.",
     "We are interested to collaborate with your company.",
     "'explore a partnership opportunity'가 외부 비즈니스 미팅에서 가장 전문적인 제안 표현이에요."),

    ("중급", "캠페인 브리핑에서 이번 캠페인의 목표를 설명하고 싶어요.",
     "The goal of this campaign is to get more users.",
     "The objective of this campaign is to drive user acquisition.",
     "The purpose of this campaign is to increase users.",
     "'objective'와 'drive user acquisition'이 마케팅 전문 용어로 더 적합해요."),

    ("중급", "프로젝트 우선순위를 논의할 때 이 건이 중요하다고 강조하고 싶어요.",
     "I think this project is more important.",
     "I'd prioritize this initiative given its potential impact.",
     "I think we should put this project first.",
     "'prioritize'와 'potential impact'이 우선순위 논의에서 가장 자연스러운 표현이에요."),

    ("중급", "개발팀에 랜딩 페이지 작업 협조를 요청하고 싶어요.",
     "Can the dev team help us with the landing page?",
     "Could we get some support from the dev team on the landing page build?",
     "Can we ask dev team to help the landing page?",
     "간접적이고 공손한 협업 요청이 팀 간 소통에서 더 자연스러워요."),

    ("중급", "바운스율 상승 원인을 데이터로 신중하게 설명하고 싶어요.",
     "The bounce rate went up because of the new design.",
     "The spike in bounce rate may be attributed to the design change.",
     "The bounce rate increased because of the new design.",
     "'may be attributed to'가 인과관계를 신중하게 표현하는 방식이에요. 데이터 발표에서는 단정보다 가능성 제시가 전문적이에요."),

    ("중급", "발표 시작 시 Q2 마케팅 전략을 소개하고 싶어요.",
     "Today I will talk about our Q2 marketing plan.",
     "Today I'd like to walk you through our Q2 marketing strategy.",
     "Today I'll present our Q2 marketing plan.",
     "'walk you through'가 청중 친화적이고 발표 시작에 가장 자연스러운 표현이에요."),

    ("중급", "동료의 피드백을 받아들이고 수정하겠다고 말하고 싶어요.",
     "OK, I will change the copy.",
     "Thanks for the input — I'll revise the copy accordingly.",
     "OK, I'll modify the copy.",
     "'Thanks for the input'으로 피드백 수용 의사를 먼저 표현하는 게 비즈니스 에티켓이에요."),

    ("중급", "현재 마감 일정이 너무 촉박하다고 재협상을 요청하고 싶어요.",
     "The deadline is too tight for us.",
     "The current timeline is challenging — could we discuss adjusting it?",
     "I don't think we can finish by the deadline.",
     "'The current timeline is challenging'이 직접 거절 없이 재협상을 요청하는 전문적인 표현이에요."),

    ("중급", "트래픽 증가 수치를 보고서에서 정확하게 표현하고 싶어요.",
     "We got 20% more traffic than last month.",
     "Traffic grew 20% month-over-month, reaching X unique visitors.",
     "Traffic increased by 20% compared to last month.",
     "'month-over-month'이 마케팅 리포트의 표준 표현이에요. 절대 수치를 함께 포함하면 더 설득력 있어요."),

    ("중급", "광고 크리에이티브가 기대만큼 성과를 내지 못했다고 피드백하고 싶어요.",
     "The ad creative is not good enough.",
     "The ad creative isn't resonating with our target audience as expected.",
     "The ad creative is not performing well.",
     "구체적인 이유(not resonating with target audience)를 포함한 피드백이 개선 방향을 명확하게 해줘요."),

    ("중급", "내일 캠페인 런칭 일정을 팀에 공지하고 싶어요.",
     "We will launch the campaign tomorrow.",
     "We're set to launch the campaign tomorrow — here's what to expect.",
     "The campaign launches tomorrow.",
     "런칭 공지에 다음 액션('here's what to expect')을 포함하면 팀 준비도를 높일 수 있어요."),

    ("중급", "경쟁사 현황을 분석하고 대응 방안을 제시하고 싶어요.",
     "Our competitor is doing better than us.",
     "Competitor X is gaining ground in our core segment — here's how we can respond.",
     "Our competitor performs better than us.",
     "구체적인 경쟁사명과 대응 방향('how we can respond')을 포함한 분석이 더 전략적이에요."),

    ("중급", "다음 캠페인의 타겟 오디언스를 구체적으로 설명하고 싶어요.",
     "We should target young people.",
     "We should focus on the 25-34 demographic who are active on Instagram.",
     "We should target the youth segment.",
     "연령대와 플랫폼을 구체화한 타겟팅 제안이 실행 가능성을 높여요."),

    ("중급", "발표 마무리에서 마케팅 성과를 총괄하여 정리하고 싶어요.",
     "In conclusion, our marketing was successful.",
     "In summary, our marketing initiatives exceeded targets in three key areas.",
     "Overall, our marketing was successful.",
     "'exceeded targets in three key areas'처럼 구체적 근거를 포함한 결론이 훨씬 설득력 있어요."),

    ("중급", "랜딩 페이지 작업에서 역할 분담을 제안하며 협업을 요청하고 싶어요.",
     "I need your help with the landing page.",
     "I was hoping we could collaborate on the landing page — I'll handle the copy if your team can take care of the design.",
     "I want your team's support for the landing page.",
     "역할 분담을 명시한 협업 요청이 실행 가능성을 높여요."),

    ("중급", "광고 예산 대비 성과가 좋았다는 것을 임원에게 설명하고 싶어요.",
     "We spent a lot, but results were good.",
     "The investment was justified — we achieved a 4x ROAS against a projected 2x.",
     "We spent much budget but we got good results.",
     "'ROAS'와 구체적인 수치 비교(4x vs projected 2x)가 예산 정당화에 핵심 요소예요."),

    ("중급", "캠페인에 문제가 발생했다고 빠르게 보고하고 싶어요.",
     "There is a problem with the campaign.",
     "We've identified an issue with the campaign that requires immediate attention.",
     "We found a problem in the campaign.",
     "'requires immediate attention'이 긴박성을 전달하면서도 전문적인 이슈 보고 표현이에요."),

    ("중급", "데이터를 근거로 영상 광고가 더 효과적이라고 제안하고 싶어요.",
     "I think it's better to use video ads.",
     "Based on our data, video ads tend to generate higher engagement in this segment.",
     "Video ads would be more effective I think.",
     "데이터 근거('Based on our data')를 포함한 의견 제시가 더 설득력 있어요."),

    ("중급", "발표를 마치며 질문이 있는지 청중에게 물어보고 싶어요.",
     "That's all I have to say.",
     "That covers everything on my end — does anyone have questions or concerns?",
     "I've finished my presentation.",
     "'Does anyone have questions or concerns?'로 마무리하면 청중 참여를 유도하는 전문적인 발표 마무리가 돼요."),

    ("중급", "광고 헤드라인을 메시지와 일치하도록 수정해달라고 요청하고 싶어요.",
     "Please change the headline of the ad.",
     "Could you revise the ad headline to better align with our messaging?",
     "Can you update the ad headline?",
     "이유('to better align with our messaging')를 포함한 수정 요청이 더 효과적이에요."),

    ("중급", "팀에 월말 마감 타임라인을 공유하고 싶어요.",
     "We plan to finish by end of month.",
     "We're targeting end of month for delivery, with a checkpoint review on the 20th.",
     "We will complete this by end of month.",
     "중간 체크포인트를 포함한 타임라인이 프로젝트 관리에 더 전문적이에요."),

    ("중급", "오늘 중으로 성과 리포트를 보내겠다고 이메일로 알리고 싶어요.",
     "I'll send the report today.",
     "I'll have the performance report in your inbox by EOD.",
     "I'll share the report today.",
     "'in your inbox by EOD'(end of day)가 마케터 이메일 커뮤니케이션의 표준 표현이에요."),

    ("중급", "여러 채널을 활용한 마케팅 전략을 소개하고 싶어요.",
     "We use many channels for marketing.",
     "We employ an omni-channel approach, including paid social, SEO, and email.",
     "We use multiple channels for marketing.",
     "'omni-channel approach'와 구체적 채널 나열이 전략 발표에서 전문적인 표현이에요."),

    ("중급", "우리 성과 지표가 업계 평균보다 높다는 것을 비교해서 말하고 싶어요.",
     "Our numbers are better than industry average.",
     "Our metrics outpace industry benchmarks by roughly 15%.",
     "Our results are better than the industry standard.",
     "'outpace industry benchmarks'와 구체적 수치(15%)가 비교 보고에 더 정확한 표현이에요."),

    ("중급", "A/B 테스트에서 B 버전이 더 높은 CTR을 달성했다고 보고하고 싶어요.",
     "Variant B got more clicks.",
     "Variant B achieved a 23% higher CTR, making it the clear winner.",
     "Variant B had more clicks than Variant A.",
     "수치(23%)와 'clear winner'로 의사결정을 돕는 표현이 테스트 리포트에 효과적이에요."),

    ("중급", "이메일로 캠페인 성과 리포트를 공식적으로 공유하고 싶어요.",
     "Here is the campaign report.",
     "Please find attached the Q2 campaign performance report.",
     "I'm attaching the campaign report.",
     "'Please find attached'가 공식 이메일에서 첨부 파일을 공유할 때 쓰는 표준 표현이에요."),

    ("중급", "광고 디자인이 너무 복잡해서 개선이 필요하다고 피드백하고 싶어요.",
     "The design is too complex.",
     "The visual is a bit cluttered — simplifying it could improve click-through rates.",
     "The design has too many elements.",
     "구체적인 개선 효과(improve CTR)를 포함한 피드백이 더 실행 가능한 방향을 제시해요."),

    ("중급", "오디언스 세분화 방식을 팀에 설명하고 싶어요.",
     "We divided our audience into groups.",
     "We segmented the audience into four distinct cohorts based on purchase intent.",
     "We separated our audience into segments.",
     "'segmented into cohorts based on purchase intent'이 마케팅 전문 용어로 더 정확한 표현이에요."),

    ("중급", "유저 퍼널 구조와 전체 전환율을 설명하고 싶어요.",
     "Users come to our website and then buy.",
     "Users move through our funnel from awareness to conversion, with a ~5% overall conversion rate.",
     "Users visit our site and eventually make purchases.",
     "퍼널 단계(awareness to conversion)와 전환율을 포함한 설명이 전문적이에요."),

    ("중급", "SEO 성과를 개선할 여지가 있다고 제안하고 싶어요.",
     "We can do better in SEO.",
     "There's significant room to improve our SEO performance, particularly around long-tail keywords.",
     "Our SEO can be improved.",
     "'significant room to improve'와 구체적 방향(long-tail keywords)이 개선 제안에 효과적이에요."),

    ("중급", "팀 전체에 프로젝트 현황을 업데이트하고 싶어요.",
     "I want to update everyone about the project.",
     "I'd like to bring everyone up to speed on where the project stands.",
     "I want to tell everyone about the project status.",
     "'bring everyone up to speed'가 팀 업데이트를 시작할 때 쓰는 자연스러운 관용 표현이에요."),

    ("중급", "캠페인 비용이 높지만 정당한 이유가 있다고 설명하고 싶어요.",
     "The cost of this campaign is too high.",
     "The cost is on the higher side, but the targeting precision justifies the spend.",
     "This campaign is more expensive than usual.",
     "단순한 비용 언급이 아니라 정당화 이유(targeting precision)를 포함한 표현이 전문적이에요."),

    ("중급", "매출 상승 원인을 분석하여 신중하게 설명하고 싶어요.",
     "Sales went up because of our campaign.",
     "The uplift in sales is likely driven by our recent campaign and the seasonal demand.",
     "Our campaign caused sales to increase.",
     "'likely driven by'로 복합 요인을 고려한 신중한 분석 표현이에요."),

    ("중급", "다음 분기 유저 성장 목표를 구체적으로 설정하고 공유하고 싶어요.",
     "Our goal for next quarter is more users.",
     "Our Q3 goal is to grow our user base by 20%, with a focus on organic acquisition.",
     "We want to get more users next quarter.",
     "수치(20%)와 방법(organic acquisition)을 포함한 목표 설정이 더 전문적이에요."),

    ("중급", "에이전시 미팅에서 캠페인 전략 방향을 바꾸자고 제안하고 싶어요.",
     "We want to change the direction of the campaign.",
     "We'd like to discuss pivoting the campaign strategy to better align with current market trends.",
     "We need to change the campaign strategy.",
     "'pivoting the strategy to align with market trends'가 전략 변경을 제안하는 세련된 표현이에요."),

    ("중급", "에이전시가 제출한 결과물이 브리프와 다르다고 피드백하고 싶어요.",
     "This isn't what we wanted.",
     "This doesn't quite match the brief — let me walk you through the key areas we'd like to revisit.",
     "This is different from what we requested.",
     "구체적인 다음 단계('walk you through key areas')를 포함한 피드백이 더 생산적이에요."),

    ("중급", "파트너에게 데이터 공유 협업을 제안하고 싶어요.",
     "We can share data with you.",
     "We'd be open to a data-sharing arrangement that could benefit both parties.",
     "We are willing to share data with you.",
     "'arrangement that benefits both parties'가 파트너십 제안에서 상호이익을 강조하는 표현이에요."),

    ("중급", "런칭 준비가 완료됐다고 팀에 확인시켜주고 싶어요.",
     "Everything is ready for the launch.",
     "We're on track for launch — all assets are approved and the team is briefed.",
     "We finished preparing for the launch.",
     "구체적인 완료 항목(assets approved, team briefed)을 포함한 확인이 더 명확해요."),

    ("중급", "캠페인 런칭 날짜가 변경됐다고 이유와 함께 공지하고 싶어요.",
     "We are changing the campaign launch date.",
     "We're pushing the launch date back to [DATE] to allow time for final QA.",
     "The launch date has been changed.",
     "변경 이유(final QA)를 포함한 공지가 팀의 이해를 높여요."),

    ("중급", "목표를 거의 달성했지만 100%는 아니었다고 보고하고 싶어요.",
     "We almost reached our goal.",
     "We came close to our target, finishing at 92% of the Q2 goal.",
     "We nearly achieved our Q2 goal.",
     "구체적 수치(92%)를 포함한 성과 보고가 더 투명하고 신뢰감을 줘요."),

    ("중급", "데이터에서 흥미로운 인사이트를 발견해 팀과 공유하고 싶어요.",
     "I found something interesting in the data.",
     "The data reveals an interesting pattern — users who engage with video content are 3x more likely to convert.",
     "I found an interesting data insight.",
     "구체적인 수치와 인사이트를 바로 제시하는 방식이 더 설득력 있어요."),

    ("중급", "이메일 채널이 가장 높은 ROI를 낸다고 보고하고 싶어요.",
     "Email marketing is working well for us.",
     "Email remains our top-performing channel, consistently driving the highest ROI.",
     "Email marketing has good performance.",
     "'top-performing channel'과 'highest ROI'를 포함한 채널 평가가 더 전문적이에요."),

    ("중급", "다음 단계 액션 아이템을 팀에 명확하게 공유하고 싶어요.",
     "Now we need to do the next things.",
     "Our next steps are to analyze the data, refine targeting, and scale the top-performing ads.",
     "Next, we have to do several tasks.",
     "구체적인 액션 아이템 나열이 'next steps' 커뮤니케이션의 표준이에요."),

    ("중급", "이번 캠페인이 전반적으로 성공적이었다고 총평하고 싶어요.",
     "Overall, the campaign was a success.",
     "Overall, the campaign exceeded expectations — hitting 110% of our KPIs and delivering strong brand lift.",
     "The campaign generally performed well.",
     "구체적인 성과 지표(110% of KPIs, brand lift)를 포함한 총평이 훨씬 설득력 있어요."),

    # ── 고급 30 ──
    ("고급", "임원에게 마케팅 전략의 정량적 성과를 보고하고 싶어요.",
     "Our marketing strategy is working.",
     "Our marketing strategy is yielding measurable results, with a 25% YoY improvement in customer acquisition cost.",
     "Our marketing efforts have been successful.",
     "YoY 지표와 구체적 수치를 포함한 임원 보고가 더 신뢰감을 줘요."),

    ("고급", "이번 분기 성과를 근거로 예산 증액을 경영진에게 요청하고 싶어요.",
     "We need more budget to grow.",
     "I'd like to make the case for a budget increase, given the strong ROAS we've demonstrated this quarter.",
     "We should have a larger budget for growth.",
     "'make the case for'가 임원진 설득에서 논리적 근거를 강조하는 표현이에요."),

    ("고급", "리텐션 중심 성장 모델로의 전략 전환을 경영진에게 발표하고 싶어요.",
     "We will change our marketing strategy.",
     "We're proposing a strategic pivot — shifting from broad acquisition to a retention-first growth model.",
     "We plan to update our marketing strategy.",
     "'strategic pivot'과 구체적 방향(retention-first)이 전략 발표의 임팩트를 높여요."),

    ("고급", "캠페인 실행 리스크를 구체적으로 언급하며 주의를 환기시키고 싶어요.",
     "This campaign might not work.",
     "There's execution risk here, particularly around creative fatigue and audience saturation.",
     "This campaign has some risks.",
     "구체적인 리스크 요인(creative fatigue, audience saturation)을 명시한 표현이 더 전문적이에요."),

    ("고급", "경쟁사 압박에 단순히 반응하는 게 아닌 선제적 대응이 필요하다고 말하고 싶어요.",
     "We need to react to what our competitors are doing.",
     "We need to develop a proactive response to the competitive pressure, rather than simply reacting.",
     "We should respond to competitor moves.",
     "'proactive response vs simply reacting'의 대비가 전략적 사고를 보여주는 표현이에요."),

    ("고급", "새 오디언스 세그먼트가 기존보다 더 나은 성과를 낼 것이라고 예측하고 싶어요.",
     "I think the new audience will perform better.",
     "My hypothesis is that the new audience segment will outperform the existing cohort by at least 15%.",
     "The new audience should do better.",
     "'My hypothesis is that'과 구체적 수치 예측이 데이터 드리븐 마케터의 표현이에요."),

    ("고급", "복합적인 결과를 강약을 구분해 균형 있게 설명하고 싶어요.",
     "The results are mixed.",
     "The results are nuanced — while we exceeded reach targets, conversion efficiency was below benchmark.",
     "We had some good and some bad results.",
     "'nuanced'와 구체적인 강/약 대비가 복합적인 결과를 전문적으로 설명하는 표현이에요."),

    ("고급", "브랜드 마케팅 투자가 장기적으로 퍼포먼스 마케팅에도 도움이 된다고 내부 설득하고 싶어요.",
     "We should focus on brand, not just performance.",
     "Investing in brand equity now will compound our performance marketing returns over time.",
     "Brand is also important, not just performance.",
     "장기적 복리 효과(compound returns)를 강조한 표현이 내부 설득에 효과적이에요."),

    ("고급", "경영진 발표에서 결과를 Q3 전략과 연결해 소개하고 싶어요.",
     "I'll show you our results today.",
     "I'll take you through the key findings and what they mean for our Q3 strategy.",
     "I'll present today's campaign results.",
     "결과를 전략과 연결하는 발표 시작이 경영진에게 더 임팩트 있어요."),

    ("고급", "저조한 성과의 캠페인을 바로 중단하지 않고 원인 분석 후 결정하자고 제안하고 싶어요.",
     "The campaign is failing. We need to stop it.",
     "Given the underperformance, I recommend pausing the campaign and conducting a root cause analysis before reallocating the budget.",
     "The campaign isn't working and we should stop it.",
     "즉각 중단이 아닌 근본 원인 분석(RCA) 이후 결정을 제안하는 표현이 더 전략적이에요."),

    ("고급", "TikTok을 신규 채널로 소규모 테스트부터 시작하자고 제안하고 싶어요.",
     "We should try TikTok.",
     "I'd recommend piloting TikTok as a growth channel, starting with a limited test budget to validate performance before scaling.",
     "We should start using TikTok for marketing.",
     "'piloting with limited test budget before scaling'이 신규 채널 도입 시 리스크 관리 전략을 보여줘요."),

    ("고급", "매출 상승의 원인을 캠페인과 외부 요인으로 구분해 분석하고 싶어요.",
     "Our campaign increased sales.",
     "While our campaign likely contributed to the sales uplift, we need to account for external factors like seasonality and the product launch.",
     "The sales increase was because of our campaign.",
     "귀인 분석에서 외부 요인을 고려하는 표현이 더 정확하고 신뢰성이 높아요."),

    ("고급", "마지막 클릭 어트리뷰션 대신 증분성 테스트 기반 채널 믹스 최적화를 제안하고 싶어요.",
     "We need to use more channels.",
     "We should optimize our channel mix based on incrementality testing rather than last-click attribution.",
     "We should diversify our marketing channels.",
     "'incrementality testing vs last-click attribution'이 고급 마케팅 측정 방법론을 이해하는 표현이에요."),

    ("고급", "팀원별 역할을 명확히 분담해 책임감과 실행 속도를 높이고 싶어요.",
     "Let's all work together on this.",
     "I'd like each team member to own a specific workstream — this will improve accountability and execution speed.",
     "Everyone should collaborate on this project.",
     "'own a specific workstream'과 'accountability' 강조가 팀 리더십 표현에 효과적이에요."),

    ("고급", "데이터 결과가 긍정적이지만 샘플 크기 한계를 솔직하게 언급하고 싶어요.",
     "Our data shows the campaign worked.",
     "The data suggests positive results, though the sample size limits our statistical confidence.",
     "The campaign data looks positive.",
     "'sample size limits statistical confidence'가 데이터 해석의 한계를 전문적으로 표현하는 방식이에요."),

    ("고급", "CPC 안정을 가정한 다음 분기 리드 성장 예측을 보고하고 싶어요.",
     "I think next quarter will be good.",
     "Based on current trends, we're projecting a 20% QoQ growth in leads, assuming stable CPC.",
     "Next quarter should show growth.",
     "구체적인 가정(stable CPC)을 포함한 예측 보고가 신뢰도를 높여요."),

    ("고급", "C레벨 업데이트에서 Q2 결과와 하반기 전략 임플리케이션을 함께 전달하고 싶어요.",
     "Here are our Q2 results.",
     "Q2 results exceeded plan, driven by email and paid social — I'll focus on what this means for our H2 strategy.",
     "Let me share our Q2 performance.",
     "임원 보고에서 결과와 전략 임플리케이션을 동시에 연결하는 표현이 효과적이에요."),

    ("고급", "신규 피처 출시에 맞춰 마케팅 메시지를 조율하자고 프로덕트팀에 제안하고 싶어요.",
     "We need the product team's help.",
     "I'd like to align with the product team to ensure our messaging is consistent with the upcoming feature release.",
     "We should work with the product team.",
     "'align to ensure messaging consistency'가 마케팅-프로덕트 협업의 구체적 목적을 잘 표현해요."),

    ("고급", "이 캠페인이 회사 OKR과 직접 연결된다는 것을 강조하고 싶어요.",
     "This campaign supports our company goals.",
     "This initiative directly supports our company OKR of increasing net revenue retention by 10%.",
     "This campaign aligns with our company objectives.",
     "'directly supports OKR'과 구체적 지표(10% NRR)로 전략적 연계를 명확히 표현해요."),

    ("고급", "전략 문서에서 타겟팅 근거와 예산 배분에 대한 검토를 요청하고 싶어요.",
     "Please check my work.",
     "I'd appreciate your review of the strategy document — particularly the targeting rationale and budget allocation.",
     "Could you review this document?",
     "검토 요청 시 구체적인 포인트(targeting rationale, budget allocation)를 명시하면 효율적이에요."),

    ("고급", "콘텐츠 마케팅 투자 증액이 장기적으로 CAC를 낮출 수 있다고 설득하고 싶어요.",
     "We should invest more in content marketing.",
     "Content marketing has a proven track record of reducing CAC over time — a 30% increase in content investment could yield long-term efficiency gains.",
     "Content marketing is worth investing in.",
     "구체적 지표(CAC)와 수치(30%), 장기 효과를 포함한 설득이 가장 임팩트 있어요."),

    ("고급", "미디어 바잉 비용을 연간 지출 규모를 근거로 협상하고 싶어요.",
     "Can we negotiate the media buy price?",
     "Given our projected annual spend with your platform, I'd like to discuss a volume-based pricing structure.",
     "Can we get a discount on media buying?",
     "'volume-based pricing structure'가 단순 할인 요청보다 전문적인 협상 표현이에요."),

    ("고급", "복합적인 결과를 채널, 오디언스, 크리에이티브별로 나누어 설명하고 싶어요.",
     "The results are complicated to explain.",
     "The results are multifaceted — I'll break down performance by channel, audience, and creative to give a full picture.",
     "The results need more detailed explanation.",
     "'multifaceted'와 구체적 분석 방식(by channel, audience, creative)을 예고하는 표현이 발표에 효과적이에요."),

    ("고급", "클릭 기반 지표에서 비즈니스 성과 중심 KPI로 측정 체계를 전환하자고 제안하고 싶어요.",
     "We're going to change how we measure success.",
     "We're evolving our measurement framework from click-based metrics to business outcome-driven KPIs.",
     "We will change our metrics and KPIs.",
     "'evolving from click-based to business outcome-driven'이 측정 체계 전환을 전략적으로 표현하는 방식이에요."),

    ("고급", "양사 오디언스를 활용한 코마케팅을 제안하고 싶어요.",
     "We can do co-marketing together.",
     "I'd like to propose a co-marketing initiative that leverages both our audiences for mutual growth.",
     "We should try co-marketing with your brand.",
     "'leverages both audiences for mutual growth'가 파트너십 가치를 구체적으로 설명하는 표현이에요."),

    ("고급", "예산 삭감 상황에서 ROI가 높은 채널에 자원을 집중하자고 제안하고 싶어요.",
     "We can't do everything with less budget.",
     "With the reduced budget, I'd recommend focusing resources on our highest-ROI channels and pausing experimental spend.",
     "We'll have to reduce our marketing activities.",
     "삭감 상황에서 자원 집중 전략을 제시하는 표현이 리더십을 보여줘요."),

    ("고급", "어트리뷰션이 어렵다는 점을 인정하면서 멀티터치 모델 도입을 제안하고 싶어요.",
     "We can't tell exactly which channel drove sales.",
     "Attribution remains a challenge — I recommend a multi-touch model to more accurately reflect the customer journey.",
     "It's hard to know which channel caused the sales.",
     "구체적 솔루션(multi-touch attribution)을 제안하는 표현이 문제 제기만 하는 것보다 전문적이에요."),

    ("고급", "내년 마케팅 연간 계획을 회사 성장 목표와 연결해 발표하고 싶어요.",
     "This is our marketing plan for next year.",
     "This is our annual marketing roadmap, aligned to our company's growth objectives and prioritized by expected ROI.",
     "I'll present next year's marketing strategy.",
     "'roadmap aligned to growth objectives, prioritized by ROI'가 연간 계획 발표에 임팩트 있는 표현이에요."),

    ("고급", "글로벌 확장 시장을 오가닉 수요 신호 데이터를 기반으로 선정하고 싶어요.",
     "We want to expand to other countries.",
     "We're evaluating international expansion, starting with markets that show organic demand signals for our product.",
     "We plan to go global with our marketing.",
     "'organic demand signals'를 기반으로 한 데이터 드리븐 글로벌 전략 표현이에요."),

    ("고급", "마케팅팀의 장기 비전을 경영진에게 임팩트 있게 전달하고 싶어요.",
     "We want to be the best marketing team.",
     "Our vision is to build a best-in-class marketing function that drives compounding growth and brand equity over the long term.",
     "We aim to become an excellent marketing team.",
     "'best-in-class', 'compounding growth', 'brand equity'를 포함한 장기 비전 표현이 가장 임팩트 있어요."),
]

def build_excel(data, job_name, filename):
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet(f"{job_name} 100문항")

    header_fmt = wb.add_format({'bold': True, 'font_color': 'white', 'bg_color': '534AB7',
                                 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'border': 1})
    meta_fmt   = wb.add_format({'bg_color': 'C0DD97', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'font_size': 10})
    correct_fmt= wb.add_format({'bg_color': 'EAF3DE', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'font_size': 10})
    other_fmt  = wb.add_format({'bg_color': 'F1EFE8', 'valign': 'vcenter', 'text_wrap': True, 'border': 1, 'font_size': 10})

    headers = ["#","난이도","직무","상황","어색한 표현 (Question)","정답 (Correct)","오답1 (Wrong)","오답2 (Wrong)","정답 이유"]
    widths  = [5, 8, 14, 35, 38, 38, 38, 38, 45]
    for c, (h, w) in enumerate(zip(headers, widths)):
        ws.write(0, c, h, header_fmt)
        ws.set_column(c, c, w)
    ws.set_row(0, 30)

    for i, (level, situation, awkward, correct, wrong2, reason) in enumerate(data, 1):
        row = [i, level, job_name, situation, awkward, correct, wrong2, "", reason]
        fmts = [meta_fmt]*4 + [other_fmt, correct_fmt, other_fmt, other_fmt, other_fmt]
        for c, (val, fmt) in enumerate(zip(row, fmts)):
            ws.write(i, c, val, fmt)
        ws.set_row(i, 70)

    wb.close()
    print(f"Saved {filename} ({len(data)} rows)")

build_excel(data, "마케터", "/Users/design_euini/Downloads/3min_biz_english_marketer_v2.xlsx")
