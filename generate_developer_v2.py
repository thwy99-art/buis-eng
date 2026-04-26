import xlsxwriter

# (난이도, 상황(한국어 문장), 어색한표현→wrong1, 정답, wrong2, 정답이유(문장형))
data = [
    # ── 코드 리뷰 (10) ──
    ("기초", "코드 리뷰에서 동료 코드에 성능 문제가 있다고 지적하고 싶어요.",
     "This code is making performance problems.",
     "This could cause performance issues.",
     "Performance is not good in this.",
     "'could cause'는 단정 짓지 않고 문제를 지적하는 가장 자연스러운 방식이에요."),

    ("기초", "코드 리뷰에서 동료에게 왜 이 방식을 선택했는지 물어보고 싶어요.",
     "Why did you do it like this?",
     "Can you explain why you chose this approach?",
     "I cannot understand this code.",
     "기술적 결정의 이유를 묻는 가장 정중하고 자연스러운 표현이에요."),

    ("기초", "코드 리뷰에서 이 부분을 리팩토링하자고 제안하고 싶어요.",
     "We must refactor this immediately.",
     "I think we should refactor this part.",
     "Refactoring is needed by us.",
     "'I think we should'는 강요 없이 제안하는 가장 부드러운 표현이에요."),

    ("기초", "코드 리뷰에서 변수 이름이 너무 모호하다고 피드백하고 싶어요.",
     "This variable name is not so good.",
     "This variable name is a bit unclear.",
     "Variable name doesn't quite make sense.",
     "'a bit unclear'가 직접 비판보다 부드럽고 구체적인 피드백 표현이에요."),

    ("기초", "코드 리뷰에서 엣지 케이스를 처리해야 한다고 말하고 싶어요.",
     "You forgot to handle edge cases.",
     "We should also handle edge cases here.",
     "Edge cases are not handled in this.",
     "'We should'로 함께 해결한다는 뉘앙스를 주는 게 협업적인 코드 리뷰 표현이에요."),

    ("중급", "코드 리뷰에서 이 로직이 더 단순화될 수 있다고 제안하고 싶어요.",
     "This logic is too much complex.",
     "This logic could be simplified a bit.",
     "This code is overly complicated to read.",
     "'could be simplified'가 개선을 부드럽게 제안하는 코드 리뷰 표현이에요. 'too much complex'는 문법 오류예요."),

    ("중급", "코드 리뷰에서 테스트 커버리지가 부족하다고 지적하고 싶어요.",
     "There is no enough test coverage here.",
     "The test coverage here looks insufficient.",
     "Tests are not enough in this section.",
     "'looks insufficient'가 테스트 부족을 객관적으로 표현하는 방식이에요. 'no enough'는 문법 오류예요."),

    ("중급", "코드 리뷰에서 이 함수가 너무 많은 역할을 한다고 지적하고 싶어요.",
     "This function is doing too much things.",
     "This function is doing too many things.",
     "This function has too much responsibilities.",
     "'too many things'가 셀 수 있는 것의 올바른 표현이에요. 'too much things'와 'too much responsibilities'는 문법 오류예요."),

    ("고급", "코드 리뷰에서 이 패턴이 향후 유지보수를 어렵게 할 수 있다고 경고하고 싶어요.",
     "This pattern will cause problems in future.",
     "This pattern could make future maintenance more difficult.",
     "This will be hard to maintain in the future.",
     "'could make future maintenance more difficult'가 구체적이고 전문적인 코드 리뷰 피드백이에요."),

    ("고급", "코드 리뷰에서 이 구현이 요구사항과 완벽히 일치하는지 확인해달라고 요청하고 싶어요.",
     "Please check if this is meeting all the requirements properly.",
     "Could you verify this aligns with all the requirements?",
     "Can you double check if requirements are met here?",
     "'verify this aligns with'가 요구사항 확인 요청에 가장 전문적인 표현이에요."),

    # ── 기술 토론 (10) ──
    ("기초", "기술 토론에서 이 라이브러리가 우리 프로젝트에 적합한지 의견을 제시하고 싶어요.",
     "I think this library is not good for us.",
     "I'm not sure this library is the right fit for our use case.",
     "This library doesn't seem good enough for us.",
     "'not the right fit for our use case'가 기술 선택에서 적합성을 평가하는 표준 표현이에요."),

    ("기초", "기술 토론에서 이 접근 방식의 장단점을 살펴보자고 제안하고 싶어요.",
     "Let's talk about the goods and bads of this approach.",
     "Let's weigh the trade-offs of this approach.",
     "Let's talk about good and bad points of this.",
     "'goods and bads'는 영어에서 쓰지 않는 표현이에요. 'weigh the trade-offs'가 기술 토론에서 장단점을 평가한다는 의미를 가장 정확하게 전달하는 표현이에요."),

    ("중급", "기술 토론에서 현재 아키텍처로는 확장하기 어렵다고 설명하고 싶어요.",
     "Our current architecture is not scalable enough.",
     "Our current architecture may struggle to scale under increased load.",
     "The architecture can't handle more traffic.",
     "'may struggle to scale under increased load'가 확장성 문제를 신중하게 표현하는 방식이에요."),

    ("중급", "기술 토론에서 기술 부채를 해결해야 한다고 주장하고 싶어요.",
     "We have too much technical debt to ignore.",
     "The technical debt here is becoming a bottleneck — I think we should prioritize addressing it.",
     "We must fix all the technical debt right now.",
     "문제의 영향('bottleneck')과 해결 제안을 함께 표현하는 것이 더 설득력 있어요."),

    ("중급", "기술 토론에서 이 솔루션이 장기적으로 더 나은 선택이라고 설명하고 싶어요.",
     "This solution is better in the long run.",
     "This solution offers better long-term maintainability, even if the initial setup takes more time.",
     "This approach is more future-proof.",
     "장기적 이점과 단기적 비용을 함께 표현하는 것이 기술 결정에서 더 설득력 있어요."),

    ("중급", "기술 토론에서 두 가지 접근 방식을 비교해 설명하고 싶어요.",
     "There are two ways we can do this.",
     "We have two main approaches here — let me walk you through the trade-offs of each.",
     "There are two options and both have good and bad points.",
     "'walk you through the trade-offs'가 기술 비교 설명에 가장 적합한 표현이에요."),

    ("고급", "기술 토론에서 마이크로서비스 전환의 복잡성을 설명하고 싶어요.",
     "Moving to microservices will be very complex.",
     "Migrating to a microservices architecture introduces significant operational complexity.",
     "Microservices architecture is too complicated for us now.",
     "'introduces significant operational complexity'가 기술 전환 비용을 전문적으로 표현하는 방식이에요."),

    ("고급", "기술 토론에서 이 API 설계가 RESTful 원칙을 따르지 않는다고 지적하고 싶어요.",
     "This API design is not following proper REST principles.",
     "This API design deviates from RESTful conventions in a few key areas.",
     "This API doesn't follow REST properly.",
     "'deviates from RESTful conventions in key areas'가 구체적이고 전문적인 표현이에요."),

    ("고급", "기술 토론에서 캐싱 전략을 도입하면 응답 속도가 크게 개선될 것이라고 제안하고 싶어요.",
     "We should add caching to make things faster.",
     "Implementing a caching layer here could significantly reduce response times.",
     "If we cache this, it will be much faster.",
     "'implementing a caching layer'와 'reduce response times'가 기술 제안에 더 정확한 표현이에요."),

    ("고급", "기술 토론에서 이 설계 결정이 시스템 전체에 영향을 줄 수 있다고 경고하고 싶어요.",
     "This decision will affect many parts of the system.",
     "This architectural decision has broad implications across the entire system — we should align the team before proceeding.",
     "This is a big decision that affects everything.",
     "영향 범위와 팀 합의의 필요성을 함께 표현하는 것이 더 전문적이에요."),

    # ── 스탠드업/스프린트 (10) ──
    ("기초", "데일리 스탠드업에서 오늘 진행할 작업을 팀에 공유하고 싶어요.",
     "I am doing API things today.",
     "I'm working on the new API integration today.",
     "My task is about API integration.",
     "현재 진행형으로 구체적인 작업을 간결하게 전달하는 게 스탠드업 최적 표현이에요."),

    ("기초", "스탠드업에서 어제 작업이 완료됐다고 공유하고 싶어요.",
     "I finished the task of yesterday.",
     "I wrapped up the authentication module yesterday.",
     "Yesterday's task is done by me.",
     "'wrapped up'이 완료를 자연스럽게 표현하는 비즈니스 표현이에요. 구체적인 작업명을 언급하는 게 좋아요."),

    ("기초", "스탠드업에서 오늘 작업이 막혀있다고 알리고 싶어요.",
     "I am blocked and can't do work today.",
     "I'm blocked on the deployment issue — I could use some help.",
     "There is a blocker and I need help.",
     "구체적으로 어디서 막혔는지와 도움 요청을 함께 표현하는 것이 스탠드업 표준이에요."),

    ("중급", "스프린트 플래닝에서 이 태스크의 예상 작업량이 너무 크다고 말하고 싶어요.",
     "This task is too big to do in one sprint.",
     "This task seems too large to fit in one sprint — should we break it down?",
     "We cannot complete this task in one sprint.",
     "문제 제기와 함께 분해 제안을 자연스럽게 연결하는 표현이에요."),

    ("중급", "스프린트 리뷰에서 이번 스프린트에서 완료한 작업을 설명하고 싶어요.",
     "We did all the things in this sprint.",
     "This sprint, we completed the user authentication flow and the dashboard redesign.",
     "We have done the work for this sprint.",
     "구체적인 완료 항목을 나열하는 것이 스프린트 리뷰의 표준 표현이에요."),

    ("중급", "스프린트 회고에서 앞으로 개선할 점을 제안하고 싶어요.",
     "We should do better next sprint.",
     "One thing we could improve next sprint is our PR review turnaround time.",
     "Next sprint we must improve many things.",
     "구체적인 개선 항목을 제시하는 것이 회고에서 더 생산적인 표현이에요."),

    ("중급", "스프린트 플래닝에서 이 기능의 우선순위를 높여야 한다고 제안하고 싶어요.",
     "I think this feature is more important than other things.",
     "I'd suggest bumping this feature up in priority given the upcoming release.",
     "This feature should have higher priority.",
     "'bumping up in priority given'이 우선순위 변경을 근거와 함께 제안하는 자연스러운 표현이에요."),

    ("고급", "스프린트 플래닝에서 이번 스프린트의 기술 부채 해소에 시간을 할당하자고 제안하고 싶어요.",
     "We should spend some time on technical debt this sprint.",
     "I'd recommend allocating some capacity this sprint for addressing our highest-priority tech debt.",
     "We must handle technical debt in this sprint.",
     "'allocating capacity for addressing tech debt'이 스프린트 플래닝에서 더 전문적인 표현이에요."),

    ("고급", "스프린트 리뷰에서 목표를 달성하지 못한 이유를 설명하고 싶어요.",
     "We couldn't finish because there were many problems.",
     "We fell short of our sprint goal due to an unexpected infrastructure outage mid-sprint.",
     "The sprint goal was not achieved because of issues.",
     "구체적인 이유(infrastructure outage)를 포함한 설명이 팀의 신뢰도를 높여요."),

    ("고급", "스프린트 플래닝에서 이 태스크의 의존성 때문에 다음 스프린트로 미뤄야 한다고 말하고 싶어요.",
     "We can't do this task because of dependencies.",
     "This task is blocked by a dependency on the data pipeline work — I'd suggest deferring it to the next sprint.",
     "This task depends on other work and should be moved to next sprint.",
     "의존성의 구체적인 내용과 대안 제시를 함께 표현하는 것이 더 전문적이에요."),

    # ── 버그 리포트 (10) ──
    ("기초", "버그를 발견했다고 팀에 알리고 싶어요.",
     "I found a bug in the code.",
     "I've spotted a bug in the checkout flow — I'll file a ticket.",
     "There is a bug that I discovered.",
     "버그 위치와 후속 조치를 함께 표현하는 것이 더 전문적인 버그 보고 방식이에요."),

    ("기초", "버그가 어떤 환경에서 재현됐는지 설명하고 싶어요.",
     "The bug happens only in some cases.",
     "This bug is reproducible only in the production environment with iOS 16.",
     "The bug can be seen in production sometimes.",
     "재현 조건(환경, OS 버전)을 구체적으로 명시하는 것이 버그 리포트의 표준이에요."),

    ("중급", "이 버그가 사용자에게 미치는 영향을 설명하고 싶어요.",
     "This bug is affecting users badly.",
     "This bug is blocking users from completing checkout, which directly impacts revenue.",
     "Users can't do things because of this bug.",
     "영향 범위와 비즈니스 임팩트를 함께 표현하는 것이 버그 우선순위 결정에 중요해요."),

    ("중급", "버그 수정이 생각보다 복잡해서 시간이 더 필요하다고 알리고 싶어요.",
     "This bug is hard to fix and needs more time.",
     "The root cause is deeper than expected — fixing this will take an additional day.",
     "Fixing this bug will take longer time than we thought.",
     "근본 원인('root cause')과 추가 소요 시간을 구체적으로 표현하는 것이 더 전문적이에요."),

    ("중급", "이 버그가 다른 기능에도 영향을 줄 수 있다고 경고하고 싶어요.",
     "This bug might also affect other features.",
     "This bug may have a wider impact — it could affect the payment and notification modules as well.",
     "Other parts of the system might also have this bug.",
     "영향받을 수 있는 모듈을 구체적으로 명시하는 것이 더 전문적인 버그 분석이에요."),

    ("중급", "버그를 임시 방편으로 처리했다고 팀에 알리고 싶어요.",
     "I did a quick fix for this bug for now.",
     "I've applied a temporary workaround — we should revisit this with a proper fix in the next sprint.",
     "I fixed it temporarily but it needs real fix later.",
     "'temporary workaround'와 후속 조치 계획을 함께 표현하는 것이 전문적인 엔지니어링 커뮤니케이션이에요."),

    ("고급", "이 버그의 근본 원인을 분석한 결과를 팀에 공유하고 싶어요.",
     "I found why this bug happened.",
     "The root cause analysis points to a race condition in the session management module.",
     "This bug was caused by problems in session management.",
     "'root cause analysis'와 기술적 원인('race condition')을 명확히 표현하는 것이 RCA 리포트의 표준이에요."),

    ("고급", "이 버그가 특정 조건에서만 발생하는 재현이 어려운 버그임을 설명하고 싶어요.",
     "This bug is difficult to find and reproduce.",
     "This is an intermittent bug — it only surfaces under high concurrency conditions.",
     "This bug doesn't happen all the time and is hard to reproduce.",
     "'intermittent bug'와 발생 조건('high concurrency')을 명시하는 것이 더 정확한 기술적 표현이에요."),

    ("고급", "버그 수정이 다른 기능에 영향을 주지 않도록 회귀 테스트가 필요하다고 말하고 싶어요.",
     "We need to test that this fix doesn't break other things.",
     "We should run regression tests to ensure this fix doesn't introduce any unintended side effects.",
     "Testing is needed to make sure the fix is safe.",
     "'regression tests'와 'unintended side effects'가 품질 관리를 강조하는 전문적인 표현이에요."),

    ("고급", "이 버그를 재현하기 위한 최소 재현 조건을 정리해달라고 요청하고 싶어요.",
     "Can you show me how to make this bug happen?",
     "Could you put together a minimal reproducible example so we can isolate the issue?",
     "Please reproduce the bug and show me the steps.",
     "'minimal reproducible example'이 디버깅 협업에서 사용하는 표준 표현이에요."),

    # ── API/협업 (10) ──
    ("기초", "API 엔드포인트 명세를 공유해달라고 요청하고 싶어요.",
     "Please send me the API documentation.",
     "Could you share the API endpoint specs when you get a chance?",
     "I need the API documentation from you.",
     "'when you get a chance'를 덧붙이면 부드럽게 요청하는 표현이 돼요."),

    ("기초", "API 응답 형식이 예상과 다르다고 백엔드 개발자에게 알리고 싶어요.",
     "The API response is not like what I expected.",
     "The API response format doesn't match what's in the documentation.",
     "API response format is wrong compared to my expectation.",
     "'doesn't match what's in the documentation'이 API 불일치를 객관적으로 표현하는 방식이에요."),

    ("중급", "API 인증 방식을 변경해야 한다고 제안하고 싶어요.",
     "We should change the authentication method of the API.",
     "I'd suggest migrating the API authentication to OAuth 2.0 for better security.",
     "The API's current auth method needs to be updated.",
     "구체적인 대안(OAuth 2.0)과 이유(better security)를 포함한 제안이 더 설득력 있어요."),

    ("중급", "API 호출이 너무 많아서 성능 문제가 생긴다고 설명하고 싶어요.",
     "We are making too many API calls and it's slow.",
     "The number of API calls is causing a performance bottleneck — we should look into batching requests.",
     "Too many API calls are making our app slow.",
     "문제와 해결 방향('batching requests')을 함께 제시하는 것이 더 전문적이에요."),

    ("중급", "API 버전 관리가 필요하다고 팀에 제안하고 싶어요.",
     "We need to do versioning for our API.",
     "I think we should introduce API versioning to maintain backward compatibility.",
     "API versioning is important and we should start doing it.",
     "'maintain backward compatibility'가 API 버전 관리의 핵심 이유를 명확히 설명하는 표현이에요."),

    ("고급", "API 설계 시 멱등성을 보장해야 한다고 설명하고 싶어요.",
     "API requests should give same results when called multiple times.",
     "We need to ensure idempotency for our POST endpoints to prevent duplicate transactions.",
     "Calling the API multiple times should not cause problems.",
     "'idempotency'와 구체적인 위험('duplicate transactions')을 명시하는 것이 전문적인 API 설계 표현이에요."),

    ("고급", "API 응답 시간이 SLA를 초과하고 있다고 보고하고 싶어요.",
     "Our API is too slow and not meeting requirements.",
     "Our API response times are consistently exceeding our 200ms SLA — we need to investigate.",
     "The API response time is not good enough.",
     "'exceeding our 200ms SLA'처럼 구체적인 수치와 기준을 명시하는 것이 더 전문적이에요."),

    ("고급", "API 설계 시 하위 호환성을 유지하면서 변경하는 방법을 제안하고 싶어요.",
     "We should change the API but not break old clients.",
     "We can introduce breaking changes in a new API version while maintaining the existing endpoints for backward compatibility.",
     "We need to update the API without affecting current users.",
     "'breaking changes in a new version'과 'backward compatibility'가 API 마이그레이션의 핵심 개념이에요."),

    ("고급", "외부 API 의존성이 서비스 안정성에 리스크가 된다고 경고하고 싶어요.",
     "Using external API is risky for our service.",
     "Our reliance on this third-party API introduces a single point of failure — we should implement a fallback.",
     "External API can cause problems if it goes down.",
     "'single point of failure'와 'fallback' 구현 필요성이 시스템 안정성을 논의할 때 핵심 표현이에요."),

    ("고급", "API 속도 제한을 고려해서 요청 전략을 재설계해야 한다고 말하고 싶어요.",
     "We need to fix our code because of API rate limits.",
     "We should redesign our request strategy to stay within the API rate limits and avoid throttling.",
     "Our API calls will be blocked if we don't handle rate limits.",
     "'redesign request strategy'와 'avoid throttling'이 API 사용 최적화를 표현하는 전문적인 방식이에요."),

    # ── 코드 설명 (10) ──
    ("기초", "이 함수가 무엇을 하는지 팀원에게 간단히 설명하고 싶어요.",
     "This function is doing the calculation of user score.",
     "This function calculates the user's score based on their activity.",
     "The purpose of this function is calculating user score.",
     "동사 중심의 간결한 설명이 코드 설명에서 가장 자연스러운 표현이에요."),

    ("기초", "이 코드가 왜 이렇게 작성됐는지 배경을 설명하고 싶어요.",
     "I wrote this code like this because of reasons.",
     "I took this approach because the original method had performance issues at scale.",
     "This code was written this way for some reasons.",
     "구체적인 이유('performance issues at scale')를 포함한 설명이 더 설득력 있어요."),

    ("중급", "이 디자인 패턴이 왜 여기에 적합한지 설명하고 싶어요.",
     "I used this pattern because it's good.",
     "I chose the Observer pattern here because we need to decouple the event source from its handlers.",
     "This design pattern is used here for decoupling purposes.",
     "패턴 이름과 구체적인 이유를 함께 제시하는 것이 기술 설명의 표준이에요."),

    ("중급", "이 코드가 아직 완성되지 않았다고 팀원에게 알리고 싶어요.",
     "This code is not finished yet.",
     "This is a work in progress — I still need to add error handling and write tests.",
     "I haven't finished writing this code.",
     "'work in progress'와 남은 작업을 구체적으로 표현하는 것이 협업에서 더 명확해요."),

    ("중급", "이 부분은 나중에 리팩토링이 필요하다고 주석으로 표시하고 싶다고 말하고 싶어요.",
     "This part needs to be changed later.",
     "I've left a TODO here — this section needs refactoring once we finalize the data model.",
     "We should refactor this code in the future.",
     "'TODO'와 구체적인 조건을 표현하는 것이 코드 주석 문화에 맞는 표현이에요."),

    ("고급", "이 비동기 처리 방식이 경쟁 조건을 방지하는 이유를 설명하고 싶어요.",
     "This async code prevents problems with timing.",
     "This async approach prevents race conditions by ensuring each operation completes before the next begins.",
     "Using async here avoids timing issues in the code.",
     "'prevents race conditions'과 동작 원리를 명확히 설명하는 것이 더 전문적이에요."),

    ("고급", "이 추상화 레이어가 왜 필요한지 팀원에게 설명하고 싶어요.",
     "We need this abstraction layer for good reasons.",
     "This abstraction layer decouples the business logic from the data access layer, making it easier to test and swap implementations.",
     "This layer separates concerns in our codebase.",
     "추상화의 구체적인 이점(testability, swappability)을 명시하는 것이 더 설득력 있어요."),

    ("고급", "이 알고리즘의 시간 복잡도를 설명하고 싶어요.",
     "This algorithm is faster than the old one.",
     "This algorithm runs in O(n log n) time, which is a significant improvement over the previous O(n²) approach.",
     "The algorithm is more efficient because it's O(n log n).",
     "이전 복잡도와 비교해서 표현하는 것이 알고리즘 설명에서 가장 명확한 방식이에요."),

    ("고급", "이 코드의 의존성 주입 패턴을 사용하는 이유를 설명하고 싶어요.",
     "We use dependency injection here because it's the right way.",
     "Dependency injection here allows us to mock dependencies in tests and swap implementations without changing the core logic.",
     "This dependency injection pattern makes our code more flexible.",
     "DI의 실질적인 이점(mocking in tests, swapping implementations)을 구체적으로 설명하는 것이 더 전문적이에요."),

    ("고급", "이 캐싱 전략이 어떻게 작동하는지 팀원에게 설명하고 싶어요.",
     "This caching works by saving data so we don't need to get it again.",
     "We cache the result at the service layer and invalidate it whenever the underlying data changes, keeping the cache fresh without unnecessary DB hits.",
     "This caching strategy stores data to reduce database queries.",
     "캐시 무효화 전략과 캐시 신선도 유지 방법을 포함한 설명이 더 전문적이에요."),

    # ── 배포/릴리즈 (10) ──
    ("기초", "배포 전에 테스트 환경에서 확인해야 한다고 말하고 싶어요.",
     "We need to test this before putting it live.",
     "We should validate this in staging before pushing to production.",
     "This should be tested in the test environment first.",
     "'validate in staging before pushing to production'이 배포 프로세스를 표현하는 표준 표현이에요."),

    ("기초", "이번 릴리즈에 포함될 변경 사항을 팀에 공유하고 싶어요.",
     "Here are the things we're releasing today.",
     "Here's a summary of what's included in this release.",
     "These are the changes in today's release.",
     "'summary of what's included in this release'가 릴리즈 노트 공유의 표준 표현이에요."),

    ("중급", "배포 중 롤백이 필요하다고 팀에 긴급히 알리고 싶어요.",
     "We have problems in production and need to go back.",
     "We're seeing critical errors in production — I'm initiating a rollback.",
     "Production is having issues so we need to rollback.",
     "'initiating a rollback'이 긴급 상황에서 롤백을 선언하는 표준 표현이에요."),

    ("중급", "피처 플래그를 사용해 점진적으로 배포하겠다고 설명하고 싶어요.",
     "We will release this feature slowly to users.",
     "We'll roll this out gradually using feature flags to monitor performance before a full release.",
     "We're going to deploy this feature step by step.",
     "'feature flags'와 'roll out gradually'가 점진적 배포를 설명하는 전문적인 표현이에요."),

    ("중급", "배포 후 모니터링 결과를 팀에 공유하고 싶어요.",
     "Everything looks okay after deployment.",
     "Post-deploy metrics look stable — error rates and latency are within normal thresholds.",
     "After the deployment, things seem to be working fine.",
     "구체적인 지표(error rates, latency)를 포함한 배포 후 보고가 더 전문적이에요."),

    ("고급", "이번 배포가 제로 다운타임으로 진행됐다고 보고하고 싶어요.",
     "We deployed without any downtime this time.",
     "The deployment completed with zero downtime using a blue-green deployment strategy.",
     "The release was done successfully with no service interruption.",
     "'zero downtime'과 'blue-green deployment'가 배포 방식을 전문적으로 설명하는 표현이에요."),

    ("고급", "이번 릴리즈의 영향을 최소화하기 위해 카나리 배포를 사용하겠다고 제안하고 싶어요.",
     "We should test this release with some users first before full rollout.",
     "I'd recommend a canary release — routing 5% of traffic to the new version before a full rollout.",
     "Let's release to a small percentage of users first to be safe.",
     "'canary release'와 구체적인 트래픽 비율(5%)을 명시하는 것이 더 전문적인 배포 전략 표현이에요."),

    ("고급", "배포 실패 원인을 분석해 팀에 공유하고 싶어요.",
     "The deployment failed because of some configuration problems.",
     "The deployment failed due to a misconfigured environment variable in the production config.",
     "There was a configuration issue that caused the deployment to fail.",
     "실패 원인을 구체적으로 명시('misconfigured environment variable')하는 것이 사후 분석의 표준이에요."),

    ("고급", "다음 배포에서 더 나은 롤백 계획이 필요하다고 제안하고 싶어요.",
     "Next time we need a better plan when things go wrong.",
     "We should define a clearer rollback plan before our next major release to reduce mean time to recovery.",
     "We need to prepare for problems in future deployments.",
     "'mean time to recovery'가 서비스 안정성 목표를 표현하는 전문 용어예요."),

    ("고급", "핫픽스를 바로 프로덕션에 배포해야 한다고 긴급 요청하고 싶어요.",
     "We need to push this fix to production right now.",
     "We need to fast-track this hotfix to production — it's directly impacting user sign-ups.",
     "This bug fix needs to go to production urgently.",
     "'fast-track this hotfix'와 비즈니스 임팩트를 함께 표현하는 것이 긴급 배포 요청에 적합해요."),

    # ── 레거시/리팩토링 (10) ──
    ("기초", "이 코드가 오래됐고 리팩토링이 필요하다고 말하고 싶어요.",
     "This code is old and needs to be changed.",
     "This code is legacy — it would benefit from a refactor.",
     "This old code should be updated soon.",
     "'legacy'와 'would benefit from a refactor'가 레거시 코드를 지적하는 전문적인 표현이에요."),

    ("중급", "레거시 코드를 건드리면 예상치 못한 문제가 생길 수 있다고 경고하고 싶어요.",
     "Be careful with this old code because it might break things.",
     "Making changes to this legacy code carries risk — we should add test coverage before touching it.",
     "This old code is fragile and changing it may cause problems.",
     "리스크와 선행 조건(test coverage)을 함께 표현하는 것이 더 전문적이에요."),

    ("중급", "이 기능을 리팩토링하면 성능이 크게 개선될 것이라고 제안하고 싶어요.",
     "Refactoring this will make it faster.",
     "Refactoring this module could yield a significant performance improvement and reduce code complexity.",
     "If we refactor this, it will be better.",
     "성능 향상과 코드 복잡도 감소라는 구체적인 이점을 표현하는 것이 더 설득력 있어요."),

    ("중급", "리팩토링 전에 테스트를 먼저 작성해야 한다고 제안하고 싶어요.",
     "We should write tests before we change this code.",
     "Before refactoring, I'd recommend adding tests to capture the current behavior.",
     "Tests need to be written before we start refactoring.",
     "'capture the current behavior'가 리팩토링 전 테스트 작성의 목적을 명확히 표현해요."),

    ("고급", "이 코드의 기술 부채를 점진적으로 해소하는 전략을 제안하고 싶어요.",
     "We should fix all the technical debt slowly.",
     "I'd recommend a strangler fig approach — gradually replacing legacy components while keeping the system functional.",
     "We need to deal with technical debt little by little.",
     "'strangler fig approach'가 레거시 시스템을 점진적으로 교체하는 패턴의 전문 용어예요."),

    ("고급", "이 레거시 모듈이 현재 시스템의 병목이 되고 있다고 설명하고 싶어요.",
     "This old module is causing problems for our whole system.",
     "This legacy module has become a bottleneck — it's limiting our ability to scale and ship new features.",
     "The legacy module is slowing down our development.",
     "'bottleneck'과 구체적인 영향(scaling, shipping features)을 표현하는 것이 더 설득력 있어요."),

    ("고급", "이 리팩토링 작업의 범위를 명확히 정의해야 한다고 말하고 싶어요.",
     "We need to know exactly what we're refactoring.",
     "Before we start, we should clearly define the scope of this refactor to avoid scope creep.",
     "The refactoring scope needs to be decided first.",
     "'scope creep'을 방지하기 위한 범위 정의의 중요성을 표현하는 전문적인 방식이에요."),

    ("고급", "이 레거시 코드를 지금 당장 리팩토링하기보다 비즈니스 가치 기준으로 우선순위를 정해야 한다고 제안하고 싶어요.",
     "We shouldn't refactor all legacy code at the same time.",
     "Rather than refactoring everything at once, I'd suggest prioritizing by business impact and migration risk.",
     "We need to prioritize which legacy code to refactor first.",
     "비즈니스 임팩트와 마이그레이션 리스크를 기준으로 우선순위를 정하는 접근 방식이 더 전략적이에요."),

    ("고급", "레거시 시스템 마이그레이션 시 데이터 정합성을 검증해야 한다고 강조하고 싶어요.",
     "We need to make sure data is correct after migration.",
     "Data integrity validation is critical during migration — we should run parallel reads to verify consistency.",
     "Data needs to be checked before and after migration.",
     "'parallel reads to verify consistency'가 마이그레이션 데이터 정합성 검증의 구체적인 방법을 표현해요."),

    ("고급", "이 레거시 API를 즉시 제거하지 말고 지원 종료 일정을 공지해야 한다고 제안하고 싶어요.",
     "We should tell people when we're going to remove this API.",
     "We should publish a deprecation timeline for this API to give clients time to migrate.",
     "We can't remove this API without warning users.",
     "'deprecation timeline'이 API 지원 종료를 전문적으로 공지하는 표준 표현이에요."),

    # ── 문서화 (10) ──
    ("기초", "README를 업데이트해야 한다고 팀에 말하고 싶어요.",
     "The README needs to be changed.",
     "The README needs to be updated to reflect the recent changes.",
     "README is outdated and should be fixed.",
     "'updated to reflect the recent changes'가 문서 업데이트 이유를 명확히 표현해요."),

    ("기초", "이 함수에 주석이 필요하다고 말하고 싶어요.",
     "This function has no comments and needs some.",
     "This function could use a brief comment explaining the expected input format.",
     "Please add comments to explain this function.",
     "주석이 필요한 이유와 내용을 구체적으로 표현하는 것이 더 유용한 피드백이에요."),

    ("중급", "API 문서가 최신 상태가 아니라고 지적하고 싶어요.",
     "The API documentation is wrong and not updated.",
     "The API documentation is out of date — several endpoints added last sprint aren't documented yet.",
     "API docs are outdated and need to be fixed.",
     "구체적으로 어떤 부분이 누락됐는지 표현하는 것이 더 명확한 피드백이에요."),

    ("중급", "이 기능의 아키텍처 결정 기록(ADR)을 작성해야 한다고 제안하고 싶어요.",
     "We should write down why we made this architectural decision.",
     "I'd suggest documenting this as an ADR so future engineers understand the context behind this decision.",
     "This architectural decision needs to be documented somewhere.",
     "'ADR(Architecture Decision Record)'이 아키텍처 결정을 문서화하는 표준 방식이에요."),

    ("중급", "온보딩 문서를 개선해서 신규 팀원이 더 빠르게 적응할 수 있도록 하자고 제안하고 싶어요.",
     "The onboarding documents are not good enough for new members.",
     "I'd like to improve our onboarding docs — it would help new engineers get up to speed faster.",
     "New team members are having trouble because documentation is poor.",
     "'get up to speed faster'가 온보딩 문서 개선의 목적을 자연스럽게 표현해요."),

    ("고급", "런북(runbook)을 작성해서 운영 이슈 대응 시간을 줄이자고 제안하고 싶어요.",
     "We need documents for when production has problems.",
     "Creating runbooks for common production issues would help reduce our mean time to resolution.",
     "We should have documentation ready for production incidents.",
     "'runbooks'와 'mean time to resolution'이 운영 문서화의 전문 용어예요."),

    ("고급", "코드 주석이 왜(Why)보다 무엇(What)을 설명하고 있다고 피드백하고 싶어요.",
     "Comments in this code are not explaining the right things.",
     "These comments explain what the code does, but they should focus on why — the code itself already shows the what.",
     "We need better comments that explain the reasoning, not just the actions.",
     "'what vs why in comments'는 코드 주석 품질을 논의하는 핵심 원칙이에요."),

    ("고급", "기술 문서가 너무 길어서 핵심 내용을 찾기 어렵다고 피드백하고 싶어요.",
     "The documentation is too long and hard to read.",
     "The documentation is quite verbose — I'd suggest adding a TL;DR section and a quick-start guide.",
     "Technical docs need to be shorter and easier to navigate.",
     "'TL;DR section'과 'quick-start guide' 제안이 문서 개선에 구체적인 방향을 제시해요."),

    ("고급", "이 코드 변경의 이유를 PR 설명에 명확히 작성해달라고 요청하고 싶어요.",
     "Please explain why you made these code changes in the PR.",
     "Could you add more context to the PR description — specifically the motivation and the trade-offs considered?",
     "The PR description needs more explanation about the reasoning.",
     "'motivation and trade-offs considered'가 PR 설명에서 가장 중요한 요소를 명확히 표현해요."),

    ("고급", "시스템 다이어그램이 현재 아키텍처를 반영하지 못하고 있다고 지적하고 싶어요.",
     "The architecture diagram is wrong and outdated.",
     "The system architecture diagram is out of sync with the current implementation — we should update it before the next planning cycle.",
     "Our architecture diagram doesn't match the actual system anymore.",
     "'out of sync with the current implementation'이 다이어그램 불일치를 표현하는 정확한 방식이에요."),

    # ── 팀 협업 (10) ──
    ("기초", "팀원에게 PR 리뷰를 부탁하고 싶어요.",
     "Please look at my PR when you have time.",
     "Could you take a look at my PR? It's ready for review.",
     "My PR needs to be reviewed by someone.",
     "'ready for review'를 명시하면 리뷰 요청 의도가 더 명확해져요."),

    ("기초", "팀원의 도움이 필요하다고 요청하고 싶어요.",
     "I need someone to help me with this problem.",
     "I'm a bit stuck on this — would anyone be available to pair on it?",
     "Can someone help me because I'm having difficulties?",
     "'pair on it'이 협업 요청을 자연스럽게 표현하는 개발자 표현이에요."),

    ("중급", "PR 리뷰에 시간이 너무 오래 걸린다고 팀에 말하고 싶어요.",
     "PR reviews are taking too long in our team.",
     "Our PR review cycle is a bit slow — would it help to set a 24-hour review SLA?",
     "We need faster PR reviews in the team.",
     "문제 제기와 함께 구체적인 해결 제안(24-hour SLA)을 포함하는 것이 더 생산적이에요."),

    ("중급", "지식 공유 세션을 만들어서 팀 전체 역량을 높이자고 제안하고 싶어요.",
     "We should share knowledge more in the team.",
     "I'd like to propose a regular knowledge-sharing session to help level up the team's understanding of our system.",
     "Team should have more knowledge sharing activities.",
     "'level up the team's understanding'이 지식 공유의 목적을 자연스럽게 표현해요."),

    ("중급", "이 작업은 내가 직접 하는 것보다 해당 전문가에게 맡기는 게 낫다고 제안하고 싶어요.",
     "I think someone else should do this task.",
     "I think this would be better handled by someone with more context on the auth service.",
     "This task should go to a person who knows this area better.",
     "'someone with more context on X'가 전문성 기반 업무 배분을 정중하게 표현하는 방식이에요."),

    ("고급", "팀의 온콜 로테이션 구조를 개선해야 한다고 제안하고 싶어요.",
     "Our on-call rotation has problems and needs improvement.",
     "Our on-call rotation is leading to burnout — I'd suggest reducing the rotation to two-week cycles and adding better alerting.",
     "The on-call system is too stressful for the team.",
     "문제(burnout)와 구체적인 해결책(two-week cycles, better alerting)을 함께 제시하는 것이 더 설득력 있어요."),

    ("고급", "팀 간 의존성이 배포 속도를 늦추고 있다고 말하고 싶어요.",
     "Too many dependencies between teams is making us slow.",
     "Our cross-team dependencies are creating bottlenecks in our release pipeline — we should discuss better decoupling.",
     "Team dependencies are slowing down our deployments.",
     "'cross-team dependencies'와 'decoupling'이 팀 간 의존성 문제를 전문적으로 표현하는 용어예요."),

    ("고급", "이번 인시던트에서 얻은 교훈을 팀 전체와 공유하고 싶어요.",
     "Let me share what we learned from this incident.",
     "I'd like to share the key learnings from this incident so we can prevent similar issues in the future.",
     "We should discuss the incident and learn from it.",
     "'key learnings'와 재발 방지 목적을 명시하는 것이 인시던트 회고의 표준 표현이에요."),

    ("고급", "팀의 코딩 컨벤션을 문서화하고 린터로 자동화하자고 제안하고 싶어요.",
     "We should agree on coding standards and make sure everyone follows them.",
     "I'd suggest formalizing our coding conventions in a shared style guide and enforcing them via linter rules.",
     "Team needs consistent coding standards that are automatically checked.",
     "'formalizing in a style guide'와 'enforcing via linter'가 코딩 컨벤션 자동화의 표준 접근 방식이에요."),

    ("고급", "팀의 기술 로드맵 계획에 시니어 엔지니어 의견이 더 반영돼야 한다고 말하고 싶어요.",
     "Senior engineers should have more say in technical decisions.",
     "I think senior engineers should be more involved in shaping the technical roadmap early in the planning cycle.",
     "The technical roadmap planning should include more senior engineer input.",
     "'involved in shaping the roadmap early'가 시니어 엔지니어 참여의 시점과 방식을 명확히 표현해요."),
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

build_excel(data, "개발자", "/Users/design_euini/Downloads/3min_biz_english_developer_v2.xlsx")
