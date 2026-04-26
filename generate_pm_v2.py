import xlsxwriter

def build_excel(data, job_name, filename):
    wb = xlsxwriter.Workbook(filename)
    ws = wb.add_worksheet(f"{job_name} 100문항")

    header_fmt = wb.add_format({'bold': True, 'bg_color': '#4472C4', 'font_color': 'white',
                                 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    meta_fmt   = wb.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
    correct_fmt= wb.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter', 'text_wrap': True,
                                 'bg_color': '#E2EFDA'})
    other_fmt  = wb.add_format({'border': 1, 'align': 'left',   'valign': 'vcenter', 'text_wrap': True})

    headers = ['#', '난이도', '직군', '상황 (한국어)', '어색한 표현', '정답', '오답1', '오답2', '정답 이유']
    col_widths = [4, 8, 10, 40, 35, 35, 35, 35, 50]

    for c, (h, w) in enumerate(zip(headers, col_widths)):
        ws.write(0, c, h, header_fmt)
        ws.set_column(c, c, w)
    ws.set_row(0, 40)

    for i, (level, situation, awkward, correct, wrong2, reason) in enumerate(data, 1):
        row = [i, level, job_name, situation, awkward, correct, wrong2, "", reason]
        fmts = [meta_fmt]*4 + [other_fmt, correct_fmt, other_fmt, other_fmt, other_fmt]
        for c, (val, fmt) in enumerate(zip(row, fmts)):
            ws.write(i, c, val, fmt)
        ws.set_row(i, 70)

    wb.close()
    print(f"Saved {filename} ({len(data)} rows)")


pm_data = [
    # ── 기초 (20) ──────────────────────────────────────────────────
    # 로드맵 발표
    ("기초", "다음 분기 제품 로드맵을 팀에 공유하고 싶어요.",
     "We will doing the new feature next quarter.",
     "We're planning to release the new feature next quarter.",
     "We will do the new feature next quarter.",
     "'will doing'은 문법적으로 틀린 표현입니다. 'be planning to' 또는 'will + 동사원형'으로 써야 합니다."),

    ("기초", "로드맵 우선순위를 설명하고 싶어요.",
     "This feature is more important as that one.",
     "This feature takes priority over that one.",
     "This feature is more important than that one because of business value.",
     "'important as'는 동등 비교이므로 우선순위 맥락에 어울리지 않습니다. 'takes priority over'가 자연스럽습니다."),

    # 스프린트 플래닝
    ("기초", "이번 스프린트 목표를 팀에 설명하고 싶어요.",
     "Our goal of this sprint is finish the login flow.",
     "Our goal for this sprint is to finish the login flow.",
     "This sprint, we need finish the login flow.",
     "'goal of'보다 'goal for'가 자연스럽고, 'is finish'는 'is to finish'로 써야 합니다."),

    ("기초", "스프린트 티켓 추정 회의를 시작하고 싶어요.",
     "Let's estimate how long tickets are taking.",
     "Let's estimate the effort for each ticket.",
     "Let's figure out how long each ticket will take.",
     "'how long tickets are taking'은 진행 중인 작업을 묻는 뉘앙스입니다. 'estimate the effort'가 추정 회의에 더 적합합니다."),

    # 요구사항 정의
    ("기초", "기능 요구사항을 개발팀에 설명하고 싶어요.",
     "The user needs to able to reset their password.",
     "The user needs to be able to reset their password.",
     "Users should have the ability to reset passwords.",
     "'to able to'는 문법 오류입니다. 'to be able to'가 올바른 표현입니다."),

    ("기초", "요구사항이 변경됐음을 팀에 알리고 싶어요.",
     "The requirements have been changed since yesterday.",
     "The requirements have changed since yesterday.",
     "We updated the requirements since our last meeting.",
     "'have been changed'는 수동태로 변경의 주체를 숨깁니다. 능동형 'have changed'가 더 자연스럽습니다."),

    # 스테이크홀더 관리
    ("기초", "스테이크홀더에게 진행 상황을 업데이트하고 싶어요.",
     "I want to give you an update about of our progress.",
     "I want to give you an update on our progress.",
     "I'd like to share a progress update with you.",
     "'about of'는 문법적으로 틀린 표현입니다. 'update on'이 올바른 표현입니다."),

    ("기초", "회의 안건을 미리 공유하고 싶어요.",
     "I will send the agenda before the meeting tomorrow.",
     "I'll send the agenda ahead of tomorrow's meeting.",
     "I'll share the agenda before we meet tomorrow.",
     "'before the meeting tomorrow'보다 'ahead of tomorrow's meeting'이 더 자연스럽고 명확합니다."),

    # 우선순위 결정
    ("기초", "어떤 기능을 먼저 개발할지 결정해야 해요.",
     "We should decide which feature to develop it first.",
     "We should decide which feature to develop first.",
     "We need to figure out which feature to prioritize.",
     "'to develop it first'에서 'it'이 불필요하게 중복됩니다. 'to develop first'가 올바릅니다."),

    ("기초", "백로그에서 가장 중요한 항목을 선택하고 싶어요.",
     "Let's pick the most important item from backlog.",
     "Let's pick the most important item from the backlog.",
     "Let's identify the highest-priority item in the backlog.",
     "'from backlog'에는 정관사 'the'가 필요합니다. 'from the backlog'가 올바릅니다."),

    # 데이터/지표
    ("기초", "이번 달 주요 지표를 팀에 공유하고 싶어요.",
     "Our key metrics for this month are look good.",
     "Our key metrics for this month look good.",
     "This month's key metrics are in good shape.",
     "'are look good'은 문법 오류입니다. 'look good'이 올바릅니다."),

    ("기초", "DAU 수치가 지난주보다 증가했음을 알리고 싶어요.",
     "DAU is increased by 15% compared to last week.",
     "DAU increased by 15% compared to last week.",
     "DAU went up 15% week over week.",
     "'is increased'는 틀린 표현입니다. 수치 변화에는 'increased'(능동) 또는 'has increased'(완료)를 씁니다."),

    # 팀 미팅/리트로
    ("기초", "회고 회의에서 잘된 점을 언급하고 싶어요.",
     "I think we did a good job in this sprint.",
     "I think we did a great job this sprint.",
     "We had a strong sprint this time around.",
     "'a good job in this sprint'보다 'a great job this sprint'이 더 자연스럽고 격려하는 표현입니다."),

    ("기초", "다음 스프린트 개선점을 제안하고 싶어요.",
     "We should to improve our communication next sprint.",
     "We should improve our communication next sprint.",
     "Let's work on improving communication in the next sprint.",
     "'should to improve'는 문법 오류입니다. 조동사 뒤에는 동사원형 'should improve'가 옵니다."),

    # 프로덕트 리뷰
    ("기초", "제품 시연을 시작하려고 해요.",
     "Let me to show you the new feature we built.",
     "Let me show you the new feature we built.",
     "I'll walk you through the new feature we built.",
     "'Let me to show'는 문법 오류입니다. 'Let me + 동사원형'이므로 'Let me show'가 맞습니다."),

    ("기초", "베타 테스트 결과를 공유하고 싶어요.",
     "The beta test results was very positive.",
     "The beta test results were very positive.",
     "We got very positive results from the beta test.",
     "'results was'는 주어-동사 불일치 오류입니다. 복수 주어에는 'were'를 사용해야 합니다."),

    # 고객/사용자 소통
    ("기초", "사용자 인터뷰 일정을 잡고 싶어요.",
     "Can we schedule a interview with some users?",
     "Can we schedule an interview with some users?",
     "Can we set up user interviews for next week?",
     "'a interview'는 모음으로 시작하는 명사 앞에 'an'을 써야 하므로 'an interview'가 올바릅니다."),

    ("기초", "사용자 피드백을 팀과 공유하고 싶어요.",
     "Users are feel confused about the onboarding flow.",
     "Users are feeling confused about the onboarding flow.",
     "Users seem confused by the onboarding flow.",
     "'are feel'은 문법 오류입니다. 'are feeling' 또는 'feel'이 올바릅니다."),

    # 출시/론칭
    ("기초", "기능 출시 일정을 공유하고 싶어요.",
     "We're planning to launch the feature on next Friday.",
     "We're planning to launch the feature this Friday.",
     "We're targeting a Friday launch for this feature.",
     "'on next Friday'는 비자연스럽습니다. 'this Friday' 또는 'next Friday' 단독으로 쓰는 것이 자연스럽습니다."),

    ("기초", "출시 후 모니터링 계획을 설명하고 싶어요.",
     "We will monitor the launch for check any issues.",
     "We will monitor the launch for any issues.",
     "We'll keep an eye on the launch and watch for issues.",
     "'for check any issues'는 문법 오류입니다. 'for any issues' 또는 'to check for any issues'가 올바릅니다."),

    # ── 중급 (50) ──────────────────────────────────────────────────
    # 로드맵 발표
    ("중급", "Q3 로드맵에서 핵심 주제가 사용자 리텐션 개선임을 발표하고 싶어요.",
     "Our main focus for Q3 road map is improving the user retention.",
     "Our main focus for Q3 is improving user retention.",
     "The central theme of our Q3 roadmap is boosting user retention.",
     "'road map'은 'roadmap'으로 붙여 쓰며, 'the user retention'에서 'the'는 불필요합니다."),

    ("중급", "로드맵 변경이 비즈니스 목표와 어떻게 연결되는지 설명하고 싶어요.",
     "This change will contribute to achieve our business goal.",
     "This change will help us achieve our business goals.",
     "This change directly supports our broader business objectives.",
     "'contribute to achieve'는 'contribute to achieving' 또는 'help achieve'로 써야 합니다."),

    ("중급", "이해관계자에게 로드맵 아이템의 근거를 설명하고 싶어요.",
     "We decided to include this because it has high impact to our users.",
     "We included this because it has a high impact on our users.",
     "We prioritized this item due to its significant impact on user experience.",
     "'impact to'는 'impact on'이 올바른 표현입니다."),

    # 스프린트 플래닝
    ("중급", "스프린트 범위가 너무 크다고 팀에 말하고 싶어요.",
     "I think we are over-committing in this sprint.",
     "I think we're overcommitting this sprint.",
     "I'm worried we've taken on too much this sprint.",
     "'are over-committing in this sprint'보다 'overcommitting this sprint'이 더 간결하고 자연스럽습니다."),

    ("중급", "스프린트 목표가 측정 가능해야 한다고 제안하고 싶어요.",
     "Our sprint goal needs to be able to measure.",
     "Our sprint goal needs to be measurable.",
     "We should make sure our sprint goal is something we can measure.",
     "'to be able to measure'는 능동 주어를 필요로 합니다. 'measurable'이 올바른 형용사입니다."),

    ("중급", "기술 부채 티켓도 이번 스프린트에 포함시켜야 한다고 주장하고 싶어요.",
     "We should include some tech debt tickets in this sprint too.",
     "We should carve out time for tech debt tickets this sprint.",
     "I'd recommend reserving capacity for a few tech debt items this sprint.",
     "'include some tickets'보다 'carve out time for'가 계획적 배려를 강조하는 더 자연스러운 표현입니다."),

    # 요구사항 정의
    ("중급", "요구사항이 아직 명확하지 않으니 더 논의가 필요하다고 말하고 싶어요.",
     "The requirements are not clear enough to start developing.",
     "The requirements aren't clear enough to move forward with development.",
     "We need more clarity on the requirements before we can start building.",
     "'start developing'보다 'move forward with development'가 더 격식체이며 PM 맥락에 적합합니다."),

    ("중급", "엣지 케이스가 명세서에 빠져있다고 지적하고 싶어요.",
     "We forgot to mention the edge case in the spec.",
     "We didn't account for the edge cases in the spec.",
     "A few edge cases seem to be missing from the spec.",
     "'forgot to mention'은 실수를 인정하는 표현이고, 'didn't account for'는 고려하지 못했음을 객관적으로 말하는 표현입니다."),

    ("중급", "범위 외 기능이 포함되었다고 팀에 알리고 싶어요.",
     "This feature is not in the scope of this sprint.",
     "This feature falls outside the scope of this sprint.",
     "This feature is out of scope for the current sprint.",
     "'not in the scope'보다 'falls outside the scope' 또는 'out of scope'가 더 자연스러운 전문 표현입니다."),

    # 스테이크홀더 관리
    ("중급", "임원에게 지연 이유를 설명하고 싶어요.",
     "We are behind schedule due to the unexpected technical issue.",
     "We're behind schedule due to an unexpected technical issue.",
     "We've hit an unexpected technical snag that's pushed back our timeline.",
     "'the unexpected technical issue'에서 불특정 이슈이므로 'an'이 맞습니다."),

    ("중급", "스테이크홀더의 새 기능 요청을 다음 분기로 미루고 싶어요.",
     "We will think about your request for next quarter.",
     "We'll revisit your request next quarter.",
     "I'll add your request to our Q3 backlog for consideration.",
     "'think about'은 모호한 표현입니다. 'revisit'이 다음 분기에 구체적으로 재검토하겠다는 의지를 전달합니다."),

    ("중급", "경영진에게 리소스 추가 요청을 하고 싶어요.",
     "We need more resource to deliver this on time.",
     "We need additional resources to deliver this on time.",
     "To hit this deadline, we'll need to bring in additional resources.",
     "'more resource'는 'more resources'(복수) 또는 'additional resources'가 자연스럽습니다."),

    # 우선순위 결정
    ("중급", "영향도와 작업량 기준으로 기능 우선순위를 정하고 싶어요.",
     "Let's prioritize by impact versus effort.",
     "Let's prioritize based on impact versus effort.",
     "I suggest we rank features by their impact-to-effort ratio.",
     "'prioritize by'도 가능하지만, 'prioritize based on'이 기준을 명시할 때 더 격식체입니다."),

    ("중급", "이 기능이 핵심 사용자 가치를 제공하지 않아 보류해야 한다고 설명하고 싶어요.",
     "This feature doesn't provide core value so we should delay it.",
     "This feature doesn't deliver core user value, so I'd recommend pushing it back.",
     "Since this doesn't address a core user need, let's defer it for now.",
     "'delay'는 단순 연기이고, 'push it back'은 일정 조정의 뉘앙스로 더 자연스럽습니다."),

    ("중급", "긴급한 버그 수정이 신규 기능보다 우선되어야 한다고 주장하고 싶어요.",
     "The bug fix is more urgent than adding new feature.",
     "The bug fix is more urgent than shipping new features.",
     "We should address the critical bug before moving on to new feature work.",
     "'adding new feature'보다 'shipping new features'가 제품 배포 맥락에 더 자연스럽습니다."),

    # 데이터/지표
    ("중급", "전환율이 목표 대비 낮다는 것을 보고하고 싶어요.",
     "Our conversion rate is lower than our target.",
     "Our conversion rate is falling short of our target.",
     "We're not hitting our conversion rate target — we're about 10% below.",
     "'lower than'은 사실 전달이고, 'falling short of'는 목표 달성 실패를 강조하는 자연스러운 표현입니다."),

    ("중급", "지표 하락의 원인을 분석해야 한다고 제안하고 싶어요.",
     "We need to analyze why our metrics have been dropped.",
     "We need to investigate why our metrics have dropped.",
     "Let's dig into what's driving the drop in our metrics.",
     "'have been dropped'는 수동태 오용입니다. 자연 하락에는 능동형 'have dropped'를 씁니다."),

    ("중급", "A/B 테스트 결과가 통계적으로 유의미하다고 설명하고 싶어요.",
     "The A/B test result is statistically meaning.",
     "The A/B test results are statistically significant.",
     "We have statistically significant results from the A/B test.",
     "'statistically meaning'은 잘못된 표현입니다. 'statistically significant'가 올바른 전문 표현입니다."),

    # 팀 미팅/리트로
    ("중급", "팀이 이번 스프린트에서 좋은 협업을 보여줬다고 칭찬하고 싶어요.",
     "Everyone worked well together in this sprint.",
     "The team showed great collaboration this sprint.",
     "I really appreciated how well everyone collaborated this sprint.",
     "'worked well together'는 자연스럽지만, 'showed great collaboration'이 성과를 강조하는 더 격식있는 표현입니다."),

    ("중급", "반복되는 커뮤니케이션 문제를 회고에서 다루고 싶어요.",
     "Communication problem happened again this sprint.",
     "We ran into communication issues again this sprint.",
     "We're seeing a recurring pattern of communication breakdowns this sprint.",
     "'problem happened'는 단순 서술이고, 'ran into issues'가 겪었음을 표현하는 더 자연스러운 구어체입니다."),

    ("중급", "회고에서 팀 프로세스 개선을 위한 액션 아이템을 정하고 싶어요.",
     "Let's make action items to improve our process.",
     "Let's define action items to improve our team process.",
     "I'd like us to walk away from this retro with clear action items.",
     "'make action items'보다 'define action items'가 더 격식있고 목적을 명확히 합니다."),

    # 프로덕트 리뷰
    ("중급", "신규 기능이 사용자 목표에 얼마나 잘 맞는지 설명하고 싶어요.",
     "This feature is very aligned with user needs.",
     "This feature closely aligns with what users are trying to accomplish.",
     "The feature maps well to our users' core jobs-to-be-done.",
     "'is very aligned'보다 'closely aligns with'가 더 능동적이고 자연스러운 표현입니다."),

    ("중급", "제품 리뷰에서 UX 문제점을 지적하고 싶어요.",
     "There are some UX problem we need to fix.",
     "There are some UX issues we need to address.",
     "I noticed a few UX pain points that need to be resolved.",
     "'some UX problem'은 복수형 오류입니다. 'UX issues' 또는 'UX problems'가 올바릅니다."),

    ("중급", "이 기능이 예상보다 더 좋은 결과를 냈다고 공유하고 싶어요.",
     "This feature performed more better than we expected.",
     "This feature performed better than expected.",
     "The feature outperformed our initial expectations.",
     "'more better'는 이중 비교급 오류입니다. 'performed better than expected'가 올바릅니다."),

    # 고객/사용자 소통
    ("중급", "사용자 인터뷰에서 발견한 주요 인사이트를 팀에 공유하고 싶어요.",
     "From user interviews, we found many interesting things.",
     "Our user interviews surfaced some key insights.",
     "We pulled a few meaningful insights from the user interviews.",
     "'found many interesting things'는 모호합니다. 'surfaced key insights'가 더 전문적이고 명확합니다."),

    ("중급", "사용자 피드백이 우리 가정을 검증했다고 말하고 싶어요.",
     "The user feedback proved our assumption was right.",
     "The user feedback validated our assumptions.",
     "User feedback confirmed that our initial assumptions were on the right track.",
     "'proved our assumption was right'도 가능하지만, 'validated our assumptions'가 UX 리서치 맥락에서 더 전문적입니다."),

    ("중급", "사용자 불만 패턴이 반복되고 있음을 설명하고 싶어요.",
     "Users are always complaining about the same issue.",
     "We're seeing a recurring pattern of user complaints.",
     "Users keep running into the same friction points.",
     "'always complaining'은 다소 부정적 뉘앙스입니다. 'recurring pattern'이 더 객관적이고 분석적입니다."),

    # 출시/론칭
    ("중급", "소프트 론칭 후 문제가 발견되면 즉시 롤백할 것이라고 말하고 싶어요.",
     "If we see any issue, we will rollback the feature immediately.",
     "If we see any issues, we'll roll back the feature immediately.",
     "We're prepared to roll back immediately if any critical issues surface post-launch.",
     "'rollback'은 명사이고, 동사로는 'roll back'(두 단어)이 올바릅니다."),

    ("중급", "피처 플래그로 점진적 출시를 계획하고 있다고 설명하고 싶어요.",
     "We plan to release it slowly using feature flag.",
     "We're planning a phased rollout using feature flags.",
     "We'll use feature flags to control the gradual rollout.",
     "'release it slowly'보다 'phased rollout'이 점진적 출시를 더 전문적으로 표현합니다."),

    ("중급", "출시 후 첫 48시간 동안 지표를 긴밀히 모니터링하겠다고 알리고 싶어요.",
     "We will closely watch the metrics after the launch.",
     "We'll keep a close eye on key metrics in the first 48 hours post-launch.",
     "Our team will be actively monitoring metrics for the 48 hours following launch.",
     "'watch the metrics after the launch'은 모호합니다. '48 hours post-launch'처럼 구체적인 시간 범위를 명시하는 것이 더 명확합니다."),

    # 추가 중급 항목
    ("중급", "경쟁사 분석 결과를 팀에 공유하고 싶어요.",
     "Our competitor is doing well in this area.",
     "Competitors are gaining ground in this space.",
     "Our competitive analysis shows rivals are moving fast here.",
     "'doing well'은 너무 막연합니다. 'gaining ground'가 경쟁 우위가 좁혀지고 있다는 더 명확한 표현입니다."),

    ("중급", "MVP 범위를 팀과 합의하고 싶어요.",
     "Let's agree what features should be in the MVP.",
     "Let's align on which features should be included in the MVP.",
     "We need to reach consensus on the MVP scope before moving forward.",
     "'agree what features'보다 'align on which features'가 협의 과정을 강조하는 더 자연스러운 표현입니다."),

    ("중급", "이 기능이 왜 핵심 사용자 문제를 해결하는지 설명하고 싶어요.",
     "This feature will solve the user's problem.",
     "This feature directly addresses a core user pain point.",
     "By solving this problem, we're tackling one of our users' biggest frustrations.",
     "'solve the user's problem'은 막연합니다. 'addresses a core user pain point'가 더 구체적이고 제품 언어로 자연스럽습니다."),

    ("중급", "팀이 제품 비전을 이해하고 있는지 확인하고 싶어요.",
     "Does everyone understand what we are trying to achieve?",
     "Is everyone aligned on our product vision?",
     "I want to make sure we're all on the same page about where we're headed.",
     "'understand what we are trying to achieve'는 길고 informal합니다. 'aligned on our product vision'이 더 간결하고 전문적입니다."),

    ("중급", "이번 스프린트에서 가장 큰 기술적 리스크를 팀에 알리고 싶어요.",
     "The biggest technical risk is the database migration.",
     "The biggest technical risk this sprint is the database migration.",
     "Our main technical concern going into this sprint is the database migration.",
     "시점을 명시하는 'this sprint'를 추가하면 더 명확해집니다."),

    ("중급", "고객 이탈률이 증가하고 있어 빠른 대응이 필요하다고 말하고 싶어요.",
     "Our churn rate is going up and we need to do something.",
     "Our churn rate is climbing, and we need to act quickly.",
     "Churn is trending upward — we need to get ahead of this fast.",
     "'going up and do something'은 모호합니다. 'climbing'과 'act quickly'가 긴박함을 더 잘 전달합니다."),

    ("중급", "다음 릴리즈에서 기술 부채를 줄이는 것이 목표라고 말하고 싶어요.",
     "We want to decrease our technical debt in next release.",
     "We're aiming to reduce technical debt in the next release.",
     "Our goal for the upcoming release is to pay down some of our technical debt.",
     "'decrease'보다 'reduce'가, 'in next release'보다 'in the next release'가 더 자연스럽습니다."),

    ("중급", "사용자 리서치 없이 가정을 검증하기 어렵다고 설명하고 싶어요.",
     "Without user research, we can't check if our assumption is correct.",
     "Without user research, it's hard to validate our assumptions.",
     "We can't confidently move forward without validating these assumptions through research.",
     "'check if our assumption is correct'은 너무 직접적입니다. 'validate our assumptions'가 더 전문적입니다."),

    ("중급", "이번 분기 OKR 달성률을 보고하고 싶어요.",
     "We achieved most of our OKR this quarter.",
     "We hit most of our OKR targets this quarter.",
     "We're tracking at about 80% completion on our Q2 OKRs.",
     "'achieved our OKR'보다 'hit our OKR targets'가 목표 달성을 더 명확히 표현합니다."),

    ("중급", "신규 기능 출시가 고객 만족도에 긍정적인 영향을 미쳤다고 말하고 싶어요.",
     "Customers are more satisfying after the new feature launch.",
     "Customer satisfaction has improved since the new feature launch.",
     "We've seen a meaningful uptick in customer satisfaction following the launch.",
     "'more satisfying'은 형용사 오용입니다. 'satisfaction has improved'가 올바른 표현입니다."),

    ("중급", "내부 테스트에서 성능 저하가 발견됐다고 알리고 싶어요.",
     "We found some performance issue during internal testing.",
     "We identified performance issues during internal testing.",
     "Internal testing surfaced some performance degradation we need to address.",
     "'some performance issue'는 복수형 오류입니다. 'performance issues'가 올바르고, 'identified'가 더 전문적입니다."),

    ("중급", "다음 분기에 모바일 앱 리뉴얼을 계획하고 있다고 말하고 싶어요.",
     "We are planning to redesign the mobile app next quarter.",
     "We're planning a mobile app redesign for next quarter.",
     "We've scoped a mobile app redesign as a key initiative for Q3.",
     "'planning to redesign'보다 'planning a redesign'이 더 자연스러운 명사구 표현입니다."),

    ("중급", "기능 요청이 현재 전략과 맞지 않는다고 설명하고 싶어요.",
     "This feature request doesn't match our current strategy.",
     "This feature request isn't aligned with our current strategy.",
     "This request falls outside our current strategic focus.",
     "'doesn't match'보다 'isn't aligned with'가 전략적 맥락에서 더 자연스럽습니다."),

    # ── 고급 (30) ──────────────────────────────────────────────────
    # 로드맵 발표
    ("고급", "장기 제품 비전과 단기 실행 계획이 어떻게 연결되는지 임원에게 설명하고 싶어요.",
     "Our long-term vision connects with short-term plan by focusing on user value.",
     "Our long-term vision is grounded in short-term bets that deliver user value.",
     "Each near-term initiative is intentionally sequenced to build toward our three-year product vision.",
     "'connects with'는 연결이 모호합니다. 'grounded in'이 단기 실행이 장기 비전의 토대임을 더 잘 표현합니다."),

    # 스프린트 플래닝
    ("고급", "스프린트 속도가 추정치와 다를 때 팀에게 원인을 설명하고 싶어요.",
     "Our velocity is different from our estimation this sprint.",
     "Our velocity deviated from our estimates this sprint.",
     "We saw a significant variance between our estimated and actual velocity this sprint.",
     "'different from estimation'은 평이합니다. 'deviated from estimates'가 속도 차이를 더 정밀하게 표현합니다."),

    # 요구사항 정의
    ("고급", "모호한 요구사항이 개발 후반에 리스크를 초래할 수 있다고 경고하고 싶어요.",
     "Vague requirements will cause problems later in development.",
     "Ambiguous requirements will introduce risk late in the development cycle.",
     "Leaving requirements underspecified now will likely surface as costly rework downstream.",
     "'cause problems'는 막연합니다. 'introduce risk late in the cycle'이 리스크 관리 언어로 더 적합합니다."),

    # 스테이크홀더 관리
    ("고급", "스테이크홀더의 기대치를 조기에 조율하지 않으면 갈등이 생길 수 있다고 설명하고 싶어요.",
     "If we don't align stakeholders early, conflict will happen.",
     "Failure to align stakeholder expectations early can lead to friction down the road.",
     "Misaligned expectations tend to create costly conflict if not addressed upfront.",
     "'conflict will happen'은 단순 예측입니다. 'lead to friction down the road'가 결과를 더 세련되게 표현합니다."),

    # 우선순위 결정
    ("고급", "사용자 영향도, 전략적 적합성, 실현 가능성 세 가지 기준으로 백로그를 점수화하겠다고 설명하고 싶어요.",
     "We'll score the backlog using impact, strategy fit, and feasibility.",
     "We'll evaluate the backlog across three dimensions: user impact, strategic fit, and feasibility.",
     "I propose we score each backlog item on impact, strategic alignment, and technical feasibility.",
     "'impact, strategy fit'보다 'user impact, strategic fit, and feasibility'처럼 세 항목을 명확히 나열하는 것이 더 정확합니다."),

    # 데이터/지표
    ("고급", "코호트 분석 결과 7일 리텐션이 신규 유저에서 낮다는 것을 발표하고 싶어요.",
     "Our cohort analysis shows new users have low retention after 7 days.",
     "Cohort analysis reveals that day-7 retention is significantly lower among new users.",
     "Our cohort data shows a sharp retention drop-off at day 7, concentrated in new user segments.",
     "'shows new users have low retention after 7 days'는 어색합니다. 'day-7 retention is lower among new users'가 더 데이터 분석 맥락에 정확합니다."),

    # 팀 미팅/리트로
    ("고급", "이번 스프린트 실패가 프로세스 문제이지 사람 문제가 아님을 강조하고 싶어요.",
     "The sprint failure was caused by the process, not people.",
     "This sprint's shortfall points to a process gap, not a people issue.",
     "The root cause here is a systemic process failure — not individual performance.",
     "'caused by the process, not people'은 직설적입니다. 'process gap, not a people issue'가 더 건설적이고 심리적 안정성을 높이는 표현입니다."),

    # 프로덕트 리뷰
    ("고급", "제품이 시장에서 차별화 포인트를 잃고 있다고 임원에게 경고하고 싶어요.",
     "Our product is losing its differentiation in the market.",
     "Our product is losing its competitive differentiation.",
     "We risk eroding our market differentiation if we don't act on this soon.",
     "'in the market'은 중복 표현입니다. 'competitive differentiation'이 더 정확하며, 'risk eroding'이 긴박함을 전달합니다."),

    # 고객/사용자 소통
    ("고급", "사용자 인터뷰에서 나온 인사이트가 기존 가정을 완전히 뒤집었다고 공유하고 싶어요.",
     "The user interviews showed our assumptions were completely wrong.",
     "The user interviews fundamentally challenged our core assumptions.",
     "Our assumptions were invalidated by the user research — we need to rethink the approach.",
     "'completely wrong'은 너무 단정적입니다. 'fundamentally challenged'가 데이터 기반으로 가정을 재검토해야 함을 더 전문적으로 표현합니다."),

    # 출시/론칭
    ("고급", "출시 전 리스크 매트릭스를 검토하고 잠재적 차단 요소를 확인하고 싶어요.",
     "Before launch, we should check if there are any risks.",
     "Before launch, let's review the risk matrix and identify any potential blockers.",
     "I'd like to walk through our pre-launch risk matrix to surface any critical blockers.",
     "'check if there are any risks'는 너무 일반적입니다. 'review the risk matrix and identify blockers'가 론칭 전 체계적인 절차를 명확히 표현합니다."),

    # 추가 고급 항목
    ("고급", "제품 지표와 비즈니스 지표 간의 상관관계를 임원진에게 설명하고 싶어요.",
     "Better product metrics lead to better business metrics.",
     "Improvements in our product metrics tend to correlate with stronger business outcomes.",
     "We've established a clear correlation between product health metrics and downstream revenue impact.",
     "'lead to better business metrics'는 인과를 단정합니다. 'correlate with'가 데이터 분석에서 더 정확한 표현입니다."),

    ("고급", "여러 팀 간 의존성을 해결하지 않으면 출시가 지연될 것임을 경고하고 싶어요.",
     "If we don't fix dependencies between teams, we'll be late.",
     "Unresolved cross-team dependencies are a critical risk to our launch timeline.",
     "Without resolving these inter-team dependencies, we're looking at a significant timeline slip.",
     "'fix dependencies'는 모호합니다. 'unresolved cross-team dependencies'가 문제의 본질을 더 정확히 표현합니다."),

    ("고급", "신규 기능이 기존 수익화 모델과 어떻게 연계되는지 설명하고 싶어요.",
     "This new feature will help our monetization.",
     "This feature is designed to support our existing monetization strategy.",
     "The feature is directly tied to our core monetization model and should drive premium conversions.",
     "'help our monetization'은 막연합니다. 'support our monetization strategy'가 전략적 연계를 더 명확히 표현합니다."),

    ("고급", "기능 개발 전에 사용자 리서치를 더 해야 한다고 팀을 설득하고 싶어요.",
     "We should do more research before we start building.",
     "We'd benefit from deeper user research before committing to a build direction.",
     "Investing in further discovery now will reduce the risk of building the wrong thing.",
     "'do more research before building'은 소극적입니다. 'deeper research before committing'이 결정 이전의 검증 필요성을 더 설득력 있게 표현합니다."),

    ("고급", "시장 환경 변화로 기존 전략을 재검토해야 한다고 경영진에게 제안하고 싶어요.",
     "The market is changing so we should change our strategy.",
     "Shifting market dynamics call for a reassessment of our current strategy.",
     "Given the evolving competitive landscape, I recommend we revisit our strategic assumptions.",
     "'market is changing so we should change'는 너무 단순합니다. 'shifting market dynamics call for a reassessment'가 더 격식있는 전략 언어입니다."),

    ("고급", "현재 제품 방향이 사용자 니즈보다 내부 편의에 맞춰져 있다고 지적하고 싶어요.",
     "I think our product is being built for our convenience, not for users.",
     "Our current product direction seems to be optimized for internal convenience rather than user needs.",
     "We've been building for operational efficiency rather than solving real user problems.",
     "'built for our convenience'는 다소 직접적입니다. 'optimized for internal convenience rather than user needs'가 더 객관적이고 설득력 있는 표현입니다."),

    ("고급", "데이터 기반 결정 문화를 팀에 정착시켜야 한다고 주장하고 싶어요.",
     "We should use data more when we make decisions.",
     "We need to embed a data-driven decision-making culture across the team.",
     "Building a culture where decisions are grounded in data, not instinct, is critical to our growth.",
     "'use data more'는 지나치게 단순합니다. 'embed a data-driven culture'가 조직 수준의 변화를 강조합니다."),

    ("고급", "출시 후 사용자 세그먼트별로 성과 차이를 분석하고 싶다고 말하고 싶어요.",
     "After launch, we'll look at how different users performed.",
     "Post-launch, we'll analyze performance variances across user segments.",
     "Our post-launch analysis will break down outcomes by user segment to surface differential impacts.",
     "'how different users performed'는 막연합니다. 'performance variances across user segments'가 분석 목적을 더 명확히 표현합니다."),

    ("고급", "기술 파트너와의 통합이 로드맵 일정에 의존성을 만든다고 설명하고 싶어요.",
     "The integration with our tech partner depends on their schedule.",
     "The partner integration introduces external dependencies that could affect our roadmap timeline.",
     "Our roadmap is exposed to schedule risk because of the dependency on the third-party integration.",
     "'depends on their schedule'은 단순 서술입니다. 'introduces external dependencies'가 리스크 관리 관점에서 더 전문적입니다."),

    ("고급", "사용자 확보보다 유지가 현재 단계에서 더 중요한 전략임을 주장하고 싶어요.",
     "Keeping users is more important than getting new users right now.",
     "At this stage, retention is a higher-leverage strategy than acquisition.",
     "Given where we are in the growth curve, optimizing for retention will yield better returns than acquisition.",
     "'keeping users is more important'은 너무 단순합니다. 'retention is a higher-leverage strategy'가 투자 대비 효과를 강조하는 더 전략적인 표현입니다."),

    # 중급 추가 7개
    ("중급", "팀원에게 명확한 업무 우선순위를 전달하고 싶어요.",
     "Please work on the most important task at first.",
     "Please prioritize the most critical task first.",
     "Let's make sure everyone is focused on the highest-priority work.",
     "'at first'는 '처음에'라는 의미로 맥락에 맞지 않습니다. 'first' 단독 또는 'prioritize'를 쓰는 것이 자연스럽습니다."),

    ("중급", "다음 기능의 성공 지표를 미리 정해두고 싶어요.",
     "We need to define success metric before we start.",
     "We need to define success metrics before we start.",
     "Let's agree on what success looks like before we kick off development.",
     "'success metric'은 복수형 'success metrics'가 일반적이며, 'agree on what success looks like'가 팀 협의 맥락에 자연스럽습니다."),

    ("중급", "이번 분기 신규 기능이 예상보다 늦어질 것 같다고 알리고 싶어요.",
     "The new feature will be delivered more late than expected.",
     "The new feature will be delivered later than expected.",
     "We're tracking behind schedule on the new feature — expect a delay of about one week.",
     "'more late'는 문법 오류입니다. 비교급은 'later'입니다."),

    ("중급", "제품 피드백 루프를 짧게 유지해야 한다고 팀에 설명하고 싶어요.",
     "We should get feedback from users faster.",
     "We should shorten our feedback loop with users.",
     "Tightening the feedback loop will help us course-correct before we invest too deeply.",
     "'get feedback faster'보다 'shorten our feedback loop'이 프로세스 개선을 더 명확히 표현합니다."),

    ("중급", "이번 릴리즈에서 성능 최적화가 UX 개선보다 우선이라고 설명하고 싶어요.",
     "Performance is more important than UX improvement for this release.",
     "For this release, performance optimization takes precedence over UX improvements.",
     "Given the current stability issues, we're prioritizing performance over UX work this cycle.",
     "'more important than'보다 'takes precedence over'가 우선순위 결정을 더 전문적으로 표현합니다."),

    ("중급", "팀 전체가 같은 제품 목표를 향하고 있는지 확인하고 싶어요.",
     "I want to check if everyone has the same product goal.",
     "I want to make sure everyone is aligned on our product goals.",
     "Let's take a moment to confirm we're all pulling in the same direction.",
     "'has the same goal'은 단순 소유 표현입니다. 'aligned on'이 팀 정렬을 더 적절히 나타냅니다."),

    ("중급", "고객 지원팀에서 올라온 이슈를 다음 스프린트에 반영하고 싶어요.",
     "We should put the customer support issues into next sprint.",
     "We should incorporate the customer support issues into the next sprint.",
     "I'd like to fold in the top customer support issues when we plan the next sprint.",
     "'put into next sprint'보다 'incorporate into the next sprint'이 더 격식체이며 자연스럽습니다."),

    # 고급 추가 10개
    ("고급", "제품 로드맵이 기술 역량 한계를 감안해 현실적으로 설계되어야 한다고 주장하고 싶어요.",
     "Our roadmap should be realistic about what we can actually build.",
     "Our roadmap needs to be calibrated against our actual engineering capacity.",
     "A credible roadmap must account for our current technical constraints and team bandwidth.",
     "'realistic about what we can build'은 일상 표현입니다. 'calibrated against engineering capacity'가 더 전문적인 계획 언어입니다."),

    ("고급", "제품 전략이 단기 매출보다 장기 사용자 가치에 집중해야 한다고 설득하고 싶어요.",
     "We should focus on user value, not just short-term revenue.",
     "Our product strategy should optimize for long-term user value over short-term revenue gains.",
     "Sacrificing long-term user value for near-term revenue will erode trust and ultimately hurt growth.",
     "'focus on user value'보다 'optimize for long-term user value over short-term revenue'가 전략 방향을 더 명확히 제시합니다."),

    ("고급", "플랫폼 확장 전에 핵심 기능의 안정성을 확보해야 한다고 임원에게 설명하고 싶어요.",
     "Before expanding the platform, we need to make the core features stable.",
     "Before scaling the platform, we need to ensure core feature stability.",
     "Expanding prematurely without solidifying the core will create compounding technical and UX debt.",
     "'make the core features stable'은 직접적이지만 평이합니다. 'ensure core feature stability'가 더 전문적입니다."),

    ("고급", "사용자 획득 채널별 LTV를 비교해서 마케팅 예산 배분을 최적화해야 한다고 제안하고 싶어요.",
     "We should compare how much value users from different channels bring.",
     "We should compare LTV across acquisition channels to optimize our marketing spend.",
     "Analyzing LTV by acquisition channel will allow us to reallocate budget toward the highest-value cohorts.",
     "'compare value from different channels'은 막연합니다. 'compare LTV across acquisition channels'가 더 정확한 분석 언어입니다."),

    ("고급", "다음 분기 전략 방향을 결정하기 전에 시장 조사를 충분히 해야 한다고 주장하고 싶어요.",
     "We should do enough market research before deciding strategy.",
     "We should conduct thorough market research before committing to a strategic direction.",
     "Locking in a strategic direction without adequate market validation increases our risk of misalignment.",
     "'do enough market research'은 단순한 표현입니다. 'conduct thorough market research before committing'이 더 전략적이고 격식있는 표현입니다."),

    ("고급", "사용자 세그먼트마다 다른 온보딩 플로우가 필요하다고 설명하고 싶어요.",
     "Different users need different onboarding flows.",
     "Different user segments require tailored onboarding experiences.",
     "A one-size-fits-all onboarding approach will underserve segments with distinct needs and contexts.",
     "'different users need different flows'는 너무 단순합니다. 'tailored onboarding experiences'가 개인화 전략을 더 명확히 표현합니다."),

    ("고급", "이 기능이 핵심 사용자층에게 불균형적으로 높은 가치를 제공한다고 설명하고 싶어요.",
     "This feature is very valuable for our core users.",
     "This feature delivers disproportionate value to our core user segment.",
     "The value concentration in our power-user segment makes this a high-leverage investment.",
     "'very valuable for core users'는 막연합니다. 'delivers disproportionate value'가 투자 대비 효과의 불균형을 더 정밀하게 표현합니다."),

    ("고급", "팀이 작은 실험을 빠르게 반복하는 문화를 만들어야 한다고 주장하고 싶어요.",
     "Our team should try small experiments more quickly.",
     "We need to build a culture of rapid, small-scale experimentation.",
     "Institutionalizing a test-and-learn mindset will help us make faster, more confident product decisions.",
     "'try small experiments more quickly'는 단순합니다. 'culture of rapid experimentation'이 조직 문화 변화를 더 명확히 강조합니다."),

    ("고급", "현재 제품 지표가 표면적 성장만 반영하고 진짜 사용자 가치는 보여주지 않는다고 지적하고 싶어요.",
     "Our metrics only show growth but not real user value.",
     "Our current metrics capture surface-level growth but don't reflect true user value.",
     "Vanity metrics may show growth, but they obscure whether we're actually solving the right problems.",
     "'only show growth but not real value'는 단순합니다. 'surface-level growth vs. true user value'가 허상 지표(vanity metrics) 문제를 더 명확히 드러냅니다."),

    ("고급", "제품-시장 적합성 검증 전에 성장에 투자하는 것은 시기상조라고 설득하고 싶어요.",
     "We shouldn't spend a lot on growth before we know our product works.",
     "Investing heavily in growth before achieving product-market fit is premature.",
     "Scaling before validating product-market fit risks amplifying a broken model rather than a working one.",
     "'spend a lot on growth before knowing it works'는 구어체입니다. 'investing in growth before product-market fit'이 더 전략적인 언어입니다."),
]

build_excel(pm_data, "PM", "/Users/design_euini/Downloads/3min_biz_english_pm_v2.xlsx")
