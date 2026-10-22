// 연결 설정
// googleClientId: 구글 드라이브 연결용 (README.md "구글 드라이브 연결하기")
// msClientId: 원노트 가져오기용 (README.md "원노트 가져오기 준비") — 비워 두면 원노트 가져오기 버튼이 안내만 보여 줘요
// msTenant: 'common' = 개인·회사·학교 계정 모두 (앱 등록의 계정 유형이 "모든 조직 디렉터리 및 개인 계정"일 때)
//           앱 등록이 "내 조직만"이면 그 조직의 디렉터리(테넌트) ID를 넣어요
// spellApi: 맞춤법 검사 서버 주소 — 홈페이지(10senter)의 OpenAI 키를 함께 씀
window.NOTE_CONFIG = {
  googleClientId: '815675421175-eg16pbjm5jvbtv1rr5ot3gda8uqt27mu.apps.googleusercontent.com',
  msClientId: '4a24cb9a-3df8-4581-a0aa-e3361ee49c10',
  msTenant: 'common',
  spellApi: 'https://10senter.vercel.app/api/spellcheck',
};
