import xlsxwriter

# (난이도, 상황(한국어 문장), 어색한표현→wrong1, 정답, wrong2, 정답이유(문장형))
data = [
    # ── 디자인 리뷰 (10) ──
    ("기초", "디자인 리뷰 중 동료에게 내 작업에 대한 피드백을 요청하고 싶어요.",
     "Can I get a feedbacks on this?",
     "Can I get feedback on this?",
     "Can I get some feedbacks on this?",
     "'feedback'은 불가산 명사라서 복수형 's'를 붙이거나 관사 'a'를 쓰면 안 돼요."),

    ("기초", "디자인 리뷰에서 내가 만든 디자인을 팀에 순서대로 소개하고 싶어요.",
     "I will show you the designs what I made.",
     "I'll walk you through the designs I made.",
     "I will show you the designs that I made.",
     "'walk you through'는 순서대로 안내할 때 쓰는 발표 필수 표현이에요. 목적격 관계대명사는 생략이 가장 자연스러워요."),

    ("기초", "이전 버전보다 이번 버전이 더 개선됐다고 설명하고 싶어요.",
     "This version is more better than before.",
     "This version is better than the previous one.",
     "This version is much better than the old one.",
     "'more better'는 이중비교급 오류예요. 'better' 자체가 비교급이라서 'more'를 붙일 필요가 없어요."),

    ("기초", "디자인 리뷰에서 상대방이 한 말의 의미를 좀 더 자세히 설명해달라고 요청하고 싶어요.",
     "Can you explain me what you mean?",
     "Can you walk me through what you mean?",
     "Can you explain to me what do you mean?",
     "'explain me'는 문법 오류예요. 'explain to me'가 맞고, 간접의문문은 평서문 어순으로 써요."),

    ("중급", "플로우 전반에 대한 의견을 동료에게 부드럽게 요청하고 싶어요.",
     "I want to get your inputs about this flow.",
     "I'd love to get your input on this flow.",
     "I want to get your input about this flow.",
     "'input'은 피드백 의미로 쓸 때 불가산 명사예요. 'on this flow'가 'about'보다 더 구체적인 표현이에요."),

    ("중급", "디자인 리뷰에서 내 디자인이 상대방이 기대했던 방향과 맞는지 확인하고 싶어요.",
     "Does this design hit the mark of what you expected?",
     "Does this design meet what you had in mind?",
     "Does this design hit what you had in mind?",
     "'meet what you had in mind'가 상대방의 기대를 확인하는 가장 자연스러운 표현이에요."),

    ("중급", "디자인에 대한 팀원의 의견을 듣고 싶다고 말하고 싶어요.",
     "I'd like to hear your thought on this.",
     "I'd like to hear your thoughts on this.",
     "I'd like to hear your thinking on this.",
     "'thoughts'는 복수형으로 써야 해요. 'your thinking'은 어색한 표현이에요."),

    ("고급", "디자인 리뷰에서 내가 내린 주요 결정들을 순서대로 설명하고 싶어요.",
     "Let me take you to the key decisions I made.",
     "Let me take you through the key decisions I made.",
     "Let me walk through the key decisions I took.",
     "'take you through'가 설명하며 안내하는 올바른 표현이에요. 'take you to'는 장소로 이동하는 뜻이에요."),

    ("고급", "이번 이터레이션이 사용자 피드백을 바탕으로 진행됐다고 설명하고 싶어요.",
     "This iteration was informed from user feedback.",
     "This iteration was informed by user feedback.",
     "This iteration was informed with user feedback.",
     "'informed by'가 올바른 전치사 조합이에요. 'informed from'과 'informed with'는 비문이에요."),

    ("고급", "리뷰어에게 날카로운 피드백이 있으면 부드럽게 공유해달라고 요청하고 싶어요.",
     "I'd appreciate it if you can share critical thoughts.",
     "I'd appreciate any critical feedback you might have.",
     "I'd appreciate it if you could share critical thoughts.",
     "'I'd appreciate + 명사'가 가장 간결하고 세련된 표현이에요. 가정법에서는 'can'이 아닌 'could'를 써요."),

    # ── 피드백 전달 (10) ──
    ("기초", "동료의 디자인에서 아이콘이 어울리지 않는다고 부드럽게 말하고 싶어요.",
     "This icon is not so good.",
     "This icon doesn't quite work here.",
     "This icon is not working well.",
     "직접 비판 대신 'doesn't quite work'처럼 완곡하게 표현하는 게 영어 비즈니스 피드백 스타일이에요."),

    ("기초", "버튼 크기가 너무 크다고 구체적으로 피드백하고 싶어요.",
     "I think this button is too much big.",
     "I think this button is a bit too large.",
     "I think this button is too much large.",
     "'too much + 형용사'는 문법 오류예요. 'too + 형용사' 구조가 올바르고 'large'가 디자인 용어로 더 적합해요."),

    ("중급", "텍스트 가독성이 떨어진다고 피드백하고 싶어요.",
     "The typography is a bit hard to see.",
     "The typography feels a bit hard to read.",
     "The typography is slightly hard to look.",
     "텍스트는 'read'가 'see'보다 적합한 표현이에요. 'hard to look'은 'look at'으로 써야 해요."),

    ("중급", "요소 간 간격이 일관성이 없다고 피드백하고 싶어요.",
     "The spacing between elements feel inconsistent.",
     "The spacing between elements feels inconsistent.",
     "The spacing between elements are feeling inconsistent.",
     "주어 'spacing'은 단수라서 'feels'가 올바른 동사예요. 'are feeling'은 수일치 오류예요."),

    ("중급", "CTA 버튼이 더 눈에 띄도록 수정해달라고 요청하고 싶어요.",
     "Can you make the CTA button more standout?",
     "Can you make the CTA button stand out more?",
     "Can you make the CTA button more outstanding?",
     "'stand out'은 동사구예요. 'make + 목적어 + 동사원형' 구조를 써야 해요."),

    ("중급", "디자인이 조금 더 다듬어져야 한다고 말하고 싶어요.",
     "I feel this needs little more polish.",
     "I feel this needs a little more polish.",
     "I feel this needs little more polishing.",
     "'a little more'가 올바른 표현이에요. 'little'만 쓰면 '거의 없다'는 반대 의미가 돼요."),

    ("고급", "비주얼 계층 구조를 개선해야 한다고 피드백하고 싶어요.",
     "The visual hierarchy is something we should address on.",
     "The visual hierarchy is something we should address.",
     "The visual hierarchy is what we should address on.",
     "'address'는 타동사라서 'on' 없이 바로 목적어를 써요. 'address on'은 비문이에요."),

    ("고급", "내비게이션의 명확성이 부족하다고 디자인 피드백을 주고 싶어요.",
     "This design lacks of clarity in the navigation.",
     "This design lacks clarity in the navigation.",
     "This design is lacking of clarity in navigation.",
     "'lack'은 타동사라서 'of' 없이 바로 목적어를 써요. 'lacks of'는 비문이에요."),

    ("고급", "이 방향이 완전히 확신이 서지 않는다고 솔직하게 말하고 싶어요.",
     "I'm not entirely convince this direction is right.",
     "I'm not entirely convinced this direction is right.",
     "I'm not entirely convincing about this direction.",
     "수동의 의미로 'convinced'(과거분사)를 써야 해요. 'convince'는 현재형, 'convincing'은 능동 의미예요."),

    ("고급", "온보딩 관련 결정을 다시 검토해볼 수 있는지 제안하고 싶어요.",
     "Could we revisit the decision we made on onboarding?",
     "Could we revisit the decision we made around onboarding?",
     "Could we revisit the decision on onboarding we made?",
     "'around onboarding'이 범위를 포괄하는 더 자연스러운 표현이에요. 관계절은 수식 명사 바로 뒤에 붙여요."),

    # ── 시안 발표 (10) ──
    ("기초", "시안 발표에서 단순하게 만드는 데 집중했다고 설명하고 싶어요.",
     "I focused to make it simple.",
     "I focused on keeping it simple.",
     "I was focusing to simplify it.",
     "'focus on + -ing'가 올바른 구조예요. 'focus to'는 비문이에요."),

    ("기초", "히스토리 확인이 필요한 유저를 위해 이 화면을 설계했다고 설명하고 싶어요.",
     "I designed this screen for the case when users want to check history.",
     "I designed this screen for users who want to check their history.",
     "I designed this screen for users want to check their history.",
     "관계절 'who want to'로 자연스럽게 연결해요. 'for the case when'은 어색한 번역투예요."),

    ("기초", "메인 컬러를 파란색으로 변경했다고 발표에서 설명하고 싶어요.",
     "The main color is changed to blue.",
     "I changed the main color to blue.",
     "The main color was changed into blue color.",
     "발표 맥락에서는 능동태 'I changed'가 주체를 명확히 해서 더 자연스러워요."),

    ("중급", "앱을 처음 열었을 때 유저가 보게 되는 화면이라고 설명하고 싶어요.",
     "This screen is for when users first open the app.",
     "This screen is what users see when they first open the app.",
     "This screen appears for when users first open the app.",
     "'what users see when'이 화면의 목적을 명확하게 설명하는 표현이에요. 'for when'은 어색해요."),

    ("중급", "새 온보딩 플로우를 단계별로 소개하고 싶어요.",
     "Let me brief you about the new onboarding flow.",
     "Let me walk you through the new onboarding flow.",
     "Let me brief you on the new onboarding flow.",
     "'walk you through'가 발표에서 단계별로 안내하는 가장 자연스러운 표현이에요."),

    ("중급", "발표 시작 시 오늘 보여줄 내용을 간략히 소개하고 싶어요.",
     "Here's what I'm going to present you today.",
     "Here's what I'll be presenting today.",
     "Here's what I am presenting to you today.",
     "'presenting today'가 간결하고 자연스러워요. 'present to you today'는 문법적이지만 길고 딱딱해요."),

    ("고급", "미적 요소와 사용성 사이에서 균형을 맞추려 했다고 설명하고 싶어요.",
     "I've tried to balance aesthetics with usability.",
     "I've tried to strike a balance between aesthetics and usability.",
     "I've tried to make a balance of aesthetics with usability.",
     "'strike a balance between A and B'가 관용 표현이에요. 'balance + 목적어'는 'between'을 쓰지 않아요."),

    ("고급", "이 디자인 패턴이 Material Design에서 영감을 받았다고 설명하고 싶어요.",
     "This design pattern was inspired from Material Design.",
     "This design pattern was inspired by Material Design.",
     "This design pattern got inspiration from Material Design.",
     "'inspired by'가 올바른 전치사 조합이에요. 'inspired from'은 비문이에요."),

    ("고급", "이 결정의 배경 논거가 노트에 정리돼 있다고 설명하고 싶어요.",
     "The rationale behind this decision are explained in the notes.",
     "The rationale behind this decision is explained in the notes.",
     "The rationale behind this decision are being explained in notes.",
     "'rationale'은 단수 명사라서 'is'가 올바른 동사예요. 'are'는 수일치 오류예요."),

    ("고급", "이번 리디자인에 유저 중심 접근 방식을 적용했다고 발표하고 싶어요.",
     "I took an user-centered approach for this redesign.",
     "I took a user-centered approach to this redesign.",
     "I took user-centered approach for this redesign.",
     "'user'는 자음 발음이라서 'a user-centered'가 맞아요. 'approach to'가 'approach for'보다 자연스러운 전치사예요."),

    # ── 개발자 협업 (10) ──
    ("기초", "개발자에게 내가 보낸 스펙 파일을 확인해달라고 요청하고 싶어요.",
     "Please check the spec file what I sent.",
     "Please check the spec file I sent.",
     "Please check the spec file that I sended.",
     "목적격 관계대명사 생략이 가장 자연스러워요. 'send'의 과거형은 'sent'예요. 'sended'는 없는 단어예요."),

    ("기초", "애니메이션에 몇 가지 문제가 있다고 개발자에게 알리고 싶어요.",
     "There is some issues with the animation.",
     "There are some issues with the animation.",
     "There have some issues with animation.",
     "'issues'는 복수 명사라서 'there are'를 써요. 'there have'는 비문이에요."),

    ("기초", "에셋을 Figma에 업로드했다고 개발자에게 알리고 싶어요.",
     "I will upload the assets on Figma.",
     "I'll upload the assets to Figma.",
     "I will upload the assets in Figma.",
     "업로드는 방향성이 있어서 'to'를 써요. 'on Figma'는 위치를 나타내는 전치사로 여기선 어색해요."),

    ("중급", "구현을 마치면 알려달라고 개발자에게 요청하고 싶어요.",
     "Can you let me know when you finish to implement this?",
     "Can you let me know when you've finished implementing this?",
     "Can you let me know when you finish implementing this?",
     "'finish + -ing'가 올바른 구조예요. 'when you've finished'가 완료 시점을 가장 자연스럽게 표현해요."),

    ("중급", "컴포넌트가 예상대로 렌더링되지 않는다고 개발자에게 전달하고 싶어요.",
     "The component is not rendering like as expected.",
     "The component isn't rendering as expected.",
     "The component is not rendering as like expected.",
     "'as expected'가 관용 표현이에요. 'like as'와 'as like'는 둘 다 비문이에요."),

    ("중급", "내가 공유한 스펙에 맞게 패딩을 조정해달라고 요청하고 싶어요.",
     "Could you align the padding to match the specs which I shared?",
     "Could you align the padding to match the specs I shared?",
     "Could you align padding to match with the specs I shared?",
     "목적격 관계대명사 생략이 자연스러워요. 'match'는 타동사라서 'with' 없이 바로 목적어를 써요."),

    ("고급", "호버 상태가 부드럽게 전환돼야 한다고 개발자에게 전달하고 싶어요.",
     "The hover state should be transitioned smoothly.",
     "The hover state should transition smoothly.",
     "The hover state needs to be transitioned in a smooth way.",
     "능동태 'should transition'이 더 간결하고 자연스러워요. 피동형 표현은 불필요하게 복잡해요."),

    ("고급", "엔지니어가 추측하지 않아도 되도록 스펙에 주석을 달겠다고 말하고 싶어요.",
     "I'll annotate the specs so engineers don't need to guess.",
     "I'll annotate the specs so engineers don't have to guess.",
     "I'll annotate the specs so that engineers don't need guessing.",
     "'don't have to'가 'don't need to'보다 자연스러운 표현이에요. 'need guessing'은 비문이에요."),

    ("고급", "서로 같은 방향으로 이해하고 있는지 개발자와 확인하고 싶어요.",
     "Let's sync up to make sure we're both in the same page.",
     "Let's sync up to make sure we're on the same page.",
     "Let's sync up to make sure we have the same page.",
     "'on the same page'가 올바른 관용 표현이에요. 'in the same page'는 전치사 오류예요."),

    ("고급", "이 인터랙션이 조금 더 다듬어져야 한다고 개발자에게 전달하고 싶어요.",
     "I think this interaction needs to be refined a bit further.",
     "I think this interaction could use a bit more refinement.",
     "I think this interaction needs refining further a bit.",
     "'could use'는 개선이 필요하다는 세련된 표현이에요. 'more further'는 이중비교급 오류예요."),

    # ── 사용자 리서치 (10) ──
    ("기초", "리서치 결과에서 앱 사용이 어렵다는 유저 피드백을 공유하고 싶어요.",
     "Users are saying that the app is difficult to use it.",
     "Users are saying the app is hard to use.",
     "Users said that the app is difficult to use it.",
     "'to use it'에서 'it'은 중복 목적어예요. 'hard to use'가 더 간결하고 자연스러운 표현이에요."),

    ("기초", "지난주에 5명의 유저와 유저 테스트를 진행했다고 보고하고 싶어요.",
     "We did user testing with 5 users last week.",
     "We ran user testing with 5 users last week.",
     "We made user testing with 5 users last week.",
     "'run user testing'이 자연스러운 콜로케이션이에요. 'an user'는 오류로 'a user'가 맞아요."),

    ("중급", "유저들이 앱 내에서 자신의 위치를 파악하지 못하는 경향이 있다고 공유하고 싶어요.",
     "Users tend to lose where they are in the app.",
     "Users tend to lose track of where they are in the app.",
     "Users often lose themselves where they are in the app.",
     "'lose track of'가 위치 파악을 못하다는 의미의 올바른 관용 표현이에요."),

    ("중급", "인터뷰가 12명의 참여자와 함께 진행됐다고 보고하고 싶어요.",
     "The interview was conducted to 12 participants.",
     "The interview was conducted with 12 participants.",
     "The interview was conducted for 12 participants.",
     "인터뷰 참여자에는 'with'를 써요. 'by'는 진행자, 'to/for'는 전치사 오류예요."),

    ("중급", "유저들이 가장 많이 언급한 페인 포인트가 내비게이션 혼란이라고 공유하고 싶어요.",
     "The pain point users mentioned most is about confusing navigation.",
     "The pain point users mentioned most is confusing navigation.",
     "The pain points users mentioned most is confusing navigation.",
     "'is + 명사구'가 깔끔한 구조예요. 'is about'은 중복이에요. 주어 'pain point'는 단수라서 'is'가 맞아요."),

    ("고급", "유저 피드백을 바탕으로 얻은 인사이트를 공유하고 싶어요.",
     "Based on user feedback, we've gained valuable insights into user behaviors.",
     "Based on user feedback, we've gained valuable insights into user behavior.",
     "Based on users feedback, we've gained valuable insights into user behavior.",
     "'behavior'는 불가산 명사라서 단수형을 써요. 'users feedback'는 소유격 오류로 'user feedback'이 맞아요."),

    ("고급", "사용성 테스트 결과에서 유저들이 특정 태스크 완료에 어려움을 겪었다고 보고하고 싶어요.",
     "The usability test has revealed that users struggled to complete this task.",
     "The usability test revealed that users struggled to complete this task.",
     "The usability test revealed that users were struggling on this task.",
     "완료된 연구 결과 보고에는 단순 과거형이 맞아요. 'struggling on'은 전치사 오류예요."),

    ("고급", "결제 화면에서 유저들이 이탈하는 패턴을 발견했다고 공유하고 싶어요.",
     "We're seeing a pattern that users drop off at the payment screen.",
     "We're seeing a pattern where users drop off at the payment screen.",
     "We're seeing a pattern which users drop off at the payment screen.",
     "장소나 상황을 나타낼 때는 'where'가 자연스러워요. 'which users drop off'는 관계사 오류예요."),

    ("고급", "리서치를 통해 발견한 핵심 유저 니즈를 이 기능이 해결한다고 설명하고 싶어요.",
     "This feature addresses a core user need we identified on research.",
     "This feature addresses a core user need we identified in research.",
     "This feature addresses core user needs which we've identified on research.",
     "'in research'가 올바른 전치사 구조예요. 'on research'는 비문이에요."),

    ("고급", "리서치 결과가 초기 가설과 일치한다고 발표하고 싶어요.",
     "The findings align well with what we've assumed earlier.",
     "The findings align well with what we hypothesized earlier.",
     "The findings align well with what we've been assuming earlier.",
     "'hypothesized'가 리서치 맥락에서 더 전문적인 표현이에요. 'been assuming earlier'는 시제 오류예요."),

    # ── 디자인 시스템 (10) ──
    ("기초", "팀에 디자인 시스템의 간격 일관성 유지를 요청하고 싶어요.",
     "We should use consistent the spacing.",
     "We should keep the spacing consistent.",
     "We should use the spacing consistently.",
     "'keep + 목적어 + 형용사'가 올바른 구조예요. 'consistent'는 형용사라서 부사 자리에 쓸 수 없어요."),

    ("기초", "이 컴포넌트가 디자인 가이드라인과 일치해야 한다고 말하고 싶어요.",
     "We need to align this with our design guideline.",
     "We need to align this with our design guidelines.",
     "We should follow the design guideline of ours.",
     "'guidelines'가 복수형으로 일반적이에요. 'of ours'는 어색한 이중소유격이에요."),

    ("중급", "이 컴포넌트가 디자인 토큰을 따르지 않는다고 지적하고 싶어요.",
     "This component doesn't follow our design token.",
     "This component doesn't follow our design tokens.",
     "This component doesn't follow the design token of ours.",
     "'design tokens'는 복수형이 일반적이에요. 'of ours'는 이중소유격 오류예요."),

    ("중급", "아이콘 크기와 스타일이 일관되어야 한다고 팀에 전달하고 싶어요.",
     "Let's make sure all icons are consistent in size and the style.",
     "Let's make sure all icons are consistent in size and style.",
     "Let's make sure all icons are consistent in size and in the style.",
     "병렬 구조에서 'size and style'로 관사 없이 통일하는 게 자연스러워요."),

    ("중급", "이 컴포넌트를 개발자 참조용으로 문서화해두어야 한다고 말하고 싶어요.",
     "We should document this component for future reference of developers.",
     "We should document this component for future developer reference.",
     "We should document this component for future references of developers.",
     "'for future developer reference'가 간결한 관용 표현이에요. 'references'는 이 맥락에서 불가산 명사예요."),

    ("고급", "이 패턴이 시스템 전반의 일관성을 깨고 있다고 지적하고 싶어요.",
     "This pattern breaks the consistency we established through the system.",
     "This pattern breaks the consistency we've established throughout the system.",
     "This pattern disrupts the consistency established by us throughout the system.",
     "'throughout'이 시스템 전체를 아우를 때 적합해요. 'through'는 단순히 통과하는 의미예요."),

    ("고급", "컴포넌트 라이브러리의 오래된 패턴을 감사해야 한다고 제안하고 싶어요.",
     "The component library needs auditing for the outdated patterns.",
     "The component library needs to be audited for outdated patterns.",
     "The component library needs to audit for outdated patterns.",
     "수동태 'needs to be audited'가 올바른 구조예요. 'needs to audit'은 라이브러리가 직접 감사하는 의미라서 오류예요."),

    ("고급", "오래된 버튼 스타일을 이번 분기 안에 폐기하자고 제안하고 싶어요.",
     "We should depreciate the old button styles by end of quarter.",
     "We should deprecate the old button styles by end of quarter.",
     "We should deprecate the old button styles by the end of this quarter.",
     "'deprecate'는 기술 용어로 폐기하다는 뜻이에요. 'depreciate'는 가치가 하락하다는 전혀 다른 단어예요."),

    ("고급", "이 문제를 디자인 부채로 플래깅하고 다음 스프린트에서 처리하자고 말하고 싶어요.",
     "I'll flag this as a design debt to be addressed in the future sprint.",
     "I'll flag this as design debt to address in a future sprint.",
     "I'll flag this as a design debt for addressing in future sprint.",
     "'design debt'는 관사 없이 쓰는 전문 용어예요. 'to address'가 'to be addressed'보다 간결하고 자연스러워요."),

    ("고급", "이 패턴이 기존 모달과 동일한 문제를 해결한다고 설명하고 싶어요.",
     "This pattern solves the same problem that our existing modal does.",
     "This pattern solves the same problem our existing modal does.",
     "This pattern solves the same problem our existing modal solved.",
     "동일 비교에서 'the same ... as'가 관용 표현이에요. 관계사를 생략하는 게 더 자연스러워요."),

    # ── 핸드오프 (10) ──
    ("기초", "개발자에게 Figma 파일 링크를 보내겠다고 말하고 싶어요.",
     "I will send you the link for Figma file.",
     "I'll send you the Figma file link.",
     "I will send to you the link of Figma.",
     "'send you + 목적어'가 기본 구조예요. 'link for'는 전치사 오류이고 'send for'는 비문이에요."),

    ("기초", "에셋을 Figma에 업로드했다고 개발팀에 알리고 싶어요.",
     "I uploaded the files on Figma.",
     "I uploaded the files to Figma.",
     "I uploaded files in Figma.",
     "업로드는 방향성이 있어서 'to'를 써요. 'on Figma'는 위치 전치사로 여기서는 어색해요."),

    ("중급", "핸드오프 시 인터랙션에 주석을 달았다고 개발자에게 알리고 싶어요.",
     "I've marked all the interactions with notes so devs know how it works.",
     "I've annotated the interactions so devs know how they work.",
     "I've noted the interactions so that devs can know how it works.",
     "'annotate'가 핸드오프 맥락에서 정확한 전문 표현이에요. 'marked with notes'는 어색한 중복이에요."),

    ("중급", "레드라인이 스펙 파일에 포함돼 있다고 말하고 싶어요.",
     "The redlines are included inside the spec file.",
     "The redlines are included in the spec file.",
     "Redlines are included into the spec file.",
     "'included in'이 가장 자연스러운 표현이에요. 'inside'는 물리적 내부를 강조하고 'into'는 움직임의 전치사예요."),

    ("중급", "이 컴포넌트가 디자인과 맞는지 개발자에게 확인해달라고 요청하고 싶어요.",
     "Can you check if this is match to the design?",
     "Can you check if this matches the design?",
     "Can you check if this is matched with the design?",
     "'matches'가 올바른 동사 형태예요. 'is match to'는 비문이고 'is matched with'는 수동태 오용이에요."),

    ("고급", "핵심 디자인 결정을 담은 핸드오프 문서를 준비했다고 말하고 싶어요.",
     "I've prepared a handoff document outlining all the design decisions.",
     "I've prepared a handoff document outlining the key design decisions.",
     "I've prepared a handoff documentation outlining all design decisions.",
     "'key design decisions'가 'all the design decisions'보다 실용적이에요. 'handoff documentation'은 어색하고 'document'이 맞아요."),

    ("고급", "스펙에 대해 추가 설명이 필요하면 알려달라고 말하고 싶어요.",
     "Let me know if you need any clarifications regarding the spec.",
     "Let me know if you need any clarification on the spec.",
     "Let me know if there's anything you need to clarify about the spec.",
     "'clarification'은 불가산 명사라서 단수형을 써요. 'on the spec'이 'regarding'보다 간결해요."),

    ("고급", "엔지니어가 쉽게 탐색할 수 있도록 Figma 파일을 구성했다고 말하고 싶어요.",
     "I've set up the Figma file so it's easily for engineers to navigate.",
     "I've set up the Figma file so it's easy for engineers to navigate.",
     "I've set up the Figma file for engineers to navigate easily.",
     "'easily'는 부사라서 형용사 자리에 올 수 없어요. 'so it's easy for ... to navigate'가 가장 명확한 표현이에요."),

    ("고급", "모든 에셋이 일관된 네이밍 컨벤션을 따르도록 정리했다고 말하고 싶어요.",
     "I've made sure all assets are named with a consistent naming convention.",
     "I've made sure all assets follow a consistent naming convention.",
     "I've made sure all assets are named consistently with the naming convention.",
     "'follow a naming convention'이 자연스러운 표현이에요. 'named with a naming convention'은 어색한 중복이에요."),

    ("고급", "구현 전에 컴포넌트 스펙을 꼭 참조해달라고 요청하고 싶어요.",
     "Please refer the component specs before starting implementation.",
     "Please refer to the component specs before starting implementation.",
     "Please reference the component specs before starting the implementation.",
     "'refer to'가 올바른 표현이에요. 'refer'는 자동사라서 반드시 'to'가 필요해요."),

    # ── 스테이크홀더 미팅 (10) ──
    ("기초", "스테이크홀더 미팅에서 디자인 방향에 대해 논의하고 싶다고 말하고 싶어요.",
     "We need to discuss about the design direction.",
     "We need to discuss the design direction.",
     "We should have a discussion about the design directions.",
     "'discuss'는 타동사라서 'about' 없이 바로 목적어를 써요. 'discuss about'은 문법 오류예요."),

    ("기초", "미팅에서 제안된 방향에 동의하지 않는다고 말하고 싶어요.",
     "I'm not agree with this direction.",
     "I don't agree with this direction.",
     "I'm not agreed with this direction.",
     "'agree'는 동사라서 'be동사 + agree' 구조는 오류예요. 'I don't agree'가 올바른 부정형이에요."),

    ("중급", "스테이크홀더 우려 사항을 기록하고 후속 조치를 취하겠다고 말하고 싶어요.",
     "Let me take note of your concern and follow back.",
     "Let me note your concern and follow up.",
     "Let me take note of your concern and get back to you.",
     "'follow up'이 올바른 관용 표현이에요. 'follow back'은 SNS 맞팔 의미예요."),

    ("중급", "지난 미팅 이후 상당한 진전이 있었다고 보고하고 싶어요.",
     "We've made significant progress since the last meet.",
     "We've made significant progress since the last meeting.",
     "We've made significant progress from the last meeting.",
     "'meeting'이 올바른 명사예요. 'meet'는 동사이고 비격식 명사예요. 'since'에 관사가 필요해요."),

    ("중급", "이 방향이 경영진이 설정한 프로덕트 비전과 일치한다고 설명하고 싶어요.",
     "This aligns with the product vision that leadership had set.",
     "This aligns with the product vision leadership has set.",
     "This aligns with the product vision that leadership has set.",
     "현재 완료 'has set'이 아직 유효한 비전을 표현해요. 관계사를 생략하면 더 자연스러워요."),

    ("고급", "최신 디자인 현황을 스테이크홀더 전체에 공유하고 싶어요.",
     "I'd like to bring everyone up to speed with the latest designs.",
     "I'd like to bring everyone up to speed on the latest designs.",
     "I'd like to speed everyone up on the latest designs.",
     "'bring up to speed on'이 관용 표현이에요. 'up to speed with'는 전치사 오류예요."),

    ("고급", "미팅 시작 전 이번 세션의 목표를 팀과 공유하고 싶어요.",
     "Before we dive in, I want to align us on the goals of this session.",
     "Before we dive in, I want to align on the goals for this session.",
     "Before we dive in, I want to align everyone around goals of this session.",
     "'align on'이 관용 표현이에요. 'align us on'에서 'us'는 불필요하고 'around goals'는 전치사 오류예요."),

    ("고급", "타임라인에 영향을 줄 수 있는 리스크를 스테이크홀더에게 알리고 싶어요.",
     "I want to flag a risk that could impact on our timeline.",
     "I want to flag a risk that could impact our timeline.",
     "I want to flag a potential risk that could impact on the timeline.",
     "'impact'는 타동사라서 'on' 없이 바로 목적어를 써요. 'impact on'은 문법 오류예요."),

    ("고급", "진행하기 전에 방향에 대한 합의를 이끌어내고 싶어요.",
     "Can we get consensus of the direction before moving forward?",
     "Can we get consensus on the direction before moving forward?",
     "Can we reach a consensus about the direction before moving forward?",
     "'consensus on'이 올바른 전치사예요. 'get consensus on'이 'reach a consensus about'보다 간결해요."),

    ("고급", "피드백에 감사를 표하고 이를 반영하겠다고 말하고 싶어요.",
     "I appreciate the feedback — let me take that on board and come back to you.",
     "I appreciate the feedback — I'll take that on board and get back to you.",
     "I appreciate the feedback — I'll take it on board and get back to you with updates.",
     "'take on board'는 수용하다는 관용 표현이에요. 'get back to you'가 'come back to you'보다 더 자연스러워요."),

    # ── 이터레이션 (10) ──
    ("기초", "피드백을 반영해 디자인을 수정하겠다고 말하고 싶어요.",
     "I will revise it base on your feedback.",
     "I'll revise it based on your feedback.",
     "I will revise it basing on your feedback.",
     "'based on'은 항상 고정된 전치사구예요. 'base on'이나 'basing on'은 오류예요."),

    ("기초", "이번 디자인이 더 단순해야 한다고 의견을 제시하고 싶어요.",
     "I think we need more simple design.",
     "I think we need a simpler design.",
     "I think the design should be more simpler.",
     "비교급 'simpler'가 자연스러워요. 'more simpler'는 이중비교급 오류예요. 관사 'a'도 필요해요."),

    ("중급", "수정된 화면을 다시 한번 검토해달라고 요청하고 싶어요.",
     "Can you give a second look to this revised screen?",
     "Can you take another look at this revised screen?",
     "Can you give another look to this revised screen?",
     "'take another look at'이 관용 표현이에요. 'give a look to'는 어색한 전치사 구조예요."),

    ("중급", "논의한 내용을 바탕으로 디자인을 업데이트했다고 말하고 싶어요.",
     "I've updated the design according to what we discussed.",
     "I've updated the design based on our discussion.",
     "I've updated the design as per the discussion we had.",
     "'based on our discussion'이 가장 자연스럽고 간결한 표현이에요. 'as per'는 지나치게 격식적이에요."),

    ("중급", "이 디자인을 완성하려면 한 라운드가 더 필요하다고 말하고 싶어요.",
     "I'll need one more round for nailing this down.",
     "I'll need one more round to nail this down.",
     "I'll need another round of iteration to nail this down.",
     "'to + 동사원형'이 목적을 나타내는 올바른 구조예요. 'for nailing'은 전치사 오류예요."),

    ("고급", "지난주 피드백 대부분을 이번 버전에 반영했다고 말하고 싶어요.",
     "I've incorporated most of the feedbacks from last week.",
     "I've incorporated most of the feedback from last week.",
     "I've incorporated most of feedbacks received last week.",
     "'feedback'은 불가산 명사예요. 'feedbacks'는 비문이고 'most of feedbacks'는 관사 누락이에요."),

    ("고급", "이번 이터레이션에서 이전에 놓쳤던 엣지 케이스를 처리했다고 말하고 싶어요.",
     "This iteration addresses the edge cases that were overlooked prior.",
     "This iteration addresses the edge cases that were previously overlooked.",
     "This iteration addresses edge cases which were overlooked before.",
     "'previously overlooked'이 가장 자연스러운 표현이에요. 'prior'를 단독 부사로 쓰면 어색해요."),

    ("고급", "최종 확정 전에 이 디자인을 다양한 각도로 검토해보고 싶다고 말하고 싶어요.",
     "I want to pressure-test this design before we call it as final.",
     "I want to pressure-test this design before we call it final.",
     "I want to pressure-test this design before calling it to final.",
     "'call it final'이 관용 표현이에요. 'call it as final'과 'calling it to final'은 비문이에요."),

    ("고급", "아이디어를 시각화하기 위해 빠르게 와이어프레임을 만들어보겠다고 말하고 싶어요.",
     "I'll take a stab to create a quick wireframe to visualize the idea.",
     "I'll take a stab at a quick wireframe to visualize the idea.",
     "I'll take a quick stab at wireframing the idea.",
     "'take a stab at + 명사/-ing'가 관용 표현이에요. 'take a stab to create'는 전치사 오류예요."),

    ("고급", "확정 전에 검토할 수 있도록 몇 가지 옵션을 준비하겠다고 말하고 싶어요.",
     "Let me put together a few options for you to check before we'll commit.",
     "Let me put together a few options for you to vet before we commit.",
     "Let me put a few options together for your vetting before we commit.",
     "시간 부사절에서 'will'은 쓰지 않아요. 'for your vetting'은 어색한 명사구예요."),

    # ── 팀 미팅 (10) ──
    ("기초", "팀 미팅에서 동료에게 의견을 물어보고 싶어요.",
     "Let me know what you think it.",
     "Let me know what you think.",
     "Let me know what do you think.",
     "'what you think'는 간접의문문으로 평서문 어순을 써요. 'think it'은 중복 목적어예요."),

    ("기초", "팀 미팅에서 프로젝트 타임라인에 대해 질문하고 싶어요.",
     "I have a question about the timeline of this project.",
     "I have a question about the project timeline.",
     "I have some question about the project timeline.",
     "'project timeline'이 자연스러운 복합 명사예요. 'some question'은 단수 명사에 어색해요."),

    ("중급", "팀 미팅에서 특정 안건을 다음 미팅으로 미루자고 제안하고 싶어요.",
     "Can we postpone this topic to the next meeting?",
     "Can we table this topic for the next meeting?",
     "Can we delay this topic until the next meeting?",
     "'table this topic'이 미루다는 비즈니스 관용 표현이에요. 'put off to'는 전치사 오류로 'put off until'이 맞아요."),

    ("중급", "유저 플로우를 내가 맡겠다고 팀에게 말하고 싶어요.",
     "I'm going to take the lead on the user flow.",
     "I'll take ownership of the user flow.",
     "I'm going to own the user flow.",
     "'take ownership'이 책임감을 강조하는 비즈니스 표현이에요. 'own'은 여기서 덜 자연스러워요."),

    ("중급", "팀 미팅 후 해당 내용을 검토하고 의견을 정리해서 공유하겠다고 말하고 싶어요.",
     "I'll have a look at this and come back with my thoughts.",
     "I'll look into this and get back to you with my thoughts.",
     "I'll look at this and come back to you with my thoughts later.",
     "'look into'가 적극적으로 검토하다는 의미예요. 'get back to you'가 'come back with'보다 자연스러운 표현이에요."),

    ("고급", "팀 미팅에서 스코프 크립 리스크가 있다고 경고하고 싶어요.",
     "I think we're in a risk of scope creep here.",
     "I think we're at risk of scope creep here.",
     "I think there's a risk of scope creep happening here.",
     "'at risk of'가 관용 표현이에요. 'in a risk of'는 비문이에요."),

    ("고급", "팀이 같은 방향으로 일하고 있는지 확인하고 싶어요.",
     "Let's make sure we're all rowing in the same direction.",
     "Let's make sure we're all pulling in the same direction.",
     "Let's make sure everyone is pulling towards the same direction.",
     "'pulling in the same direction'이 관용 표현이에요. 'rowing'은 비표준이고 'towards the same direction'은 전치사 중복이에요."),

    ("고급", "과도하게 복잡하게 접근하지 말고 집중하자고 팀에 말하고 싶어요.",
     "I don't want us to boil the ocean here — let's keep focused.",
     "I don't want us to boil the ocean here — let's stay focused.",
     "I don't want to boil the ocean here — let's all stay focused.",
     "'boil the ocean'은 과도하게 복잡하게 하다는 관용 표현이에요. 'stay focused'가 'keep focused'보다 자연스러워요."),

    ("고급", "지금 한 발 물러서서 큰 그림을 보자고 팀에게 제안하고 싶어요.",
     "This is a good opportunity to step back and think about the bigger picture.",
     "This is a good opportunity to step back and look at the bigger picture.",
     "This is a good opportunity to take a step back and consider the bigger picture.",
     "'look at the bigger picture'가 관용 표현이에요. 'think about'도 맞지만 'look at'이 더 자연스러워요."),

    ("고급", "단순함이 항상 최선이라는 가정에 의문을 제기하고 싶어요.",
     "I'd like to challenge the assumption where simplicity always wins.",
     "I'd like to challenge the assumption that simplicity always wins.",
     "I'd like to question the assumption of simplicity always winning.",
     "동격절에는 'that'을 써요. 'where'는 장소나 상황절에 쓰이고 'assumption of ... winning'은 어색한 명사구예요."),
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

build_excel(data, "프로덕트 디자이너", "/Users/design_euini/Downloads/3min_biz_english_designer_v2.xlsx")
