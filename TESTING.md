# 테스트 프로세스

## 코드 수정 후 테스트 순서

### 1. 로컬 빠른 확인 (UI/기능)
- VS Code에서 `index.html` 수정
- `Go Live` 버튼 → `localhost:5500`
- 폰에서: `http://192.168.219.102:5500` (같은 와이파이)

### 2. 실제 앱으로 테스트 (Capacitor)
```bash
# index.html 수정 후 항상 이 순서로
cp ~/Desktop/buis-eng/index.html ~/Desktop/buis-eng/www/index.html
cd ~/Desktop/buis-eng && npx cap sync ios
npx cap open ios
# Xcode에서 ▶ 눌러서 폰에 설치
```

### 3. 배포 (GitHub Pages)
- Claude에게 "PR 올려줘" 요청
- GitHub에서 PR 머지 → 자동 배포
- 배포 URL: https://thwy99-art.github.io/buis-eng/

---

## 환경 정보
- 배포: GitHub Pages (main 브랜치)
- 앱 ID: `com.fluvo.app`
- Supabase 리다이렉트 URL: `com.fluvo.app://login-callback`, `https://thwy99-art.github.io/buis-eng`
- 로컬 IP: `192.168.219.102` (와이파이 바뀌면 `ipconfig getifaddr en1` 로 재확인)
