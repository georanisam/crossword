# 선생님 페이지 진짜로 잠그기 (Apps Script 5분 작업)

`index.html`에 넣은 비밀번호 화면은 **눈에 보이는 1차 잠금**입니다.
개발자도구를 쓸 줄 아는 사람은 화면을 지우고 볼 수 있고, 무엇보다
`student.html` 소스에 Apps Script 주소가 그대로 있어서 주소만 알면
`getSubmissions`(제출 기록 전체), `deleteQuiz`(퀴즈 삭제) 같은 요청을 직접 보낼 수 있습니다.

**그걸 막는 건 아래 작업뿐입니다.** 한 번만 해두면 끝입니다.

---

## 1. Apps Script 편집기 열기

퀴즈 데이터가 들어 있는 구글 시트 → 메뉴 **확장 프로그램 → Apps Script**

## 2. `doPost` 맨 앞에 검사 넣기

`function doPost(e) {` 로 시작하는 함수를 찾습니다.
그 안에서 요청 내용을 읽는 줄(보통 `JSON.parse(e.postData.contents)`)과
`action` 을 꺼내는 줄 **바로 다음에** 아래 다섯 줄을 붙여넣습니다.

```javascript
  // ▼▼ 선생님 전용 동작 보호 ▼▼
  var OPEN_ACTIONS = ['getQuizForPlay', 'submitAnswer'];   // 학생이 쓰는 것만 열어둔다
  if (OPEN_ACTIONS.indexOf(action) === -1 &&
      String(body.pw || '') !== PropertiesService.getScriptProperties().getProperty('TEACHER_PW')) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: '선생님 비밀번호가 필요합니다.' }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  // ▲▲ 여기까지 ▲▲
```

> `body` 와 `action` 은 원래 코드에서 쓰던 변수 이름에 맞춰 주세요.
> 예를 들어 원래 코드가 `var req = JSON.parse(e.postData.contents);` 라면
> `body.pw` 를 `req.pw` 로 바꾸면 됩니다.

## 3. 비밀번호를 스크립트 속성에 저장

Apps Script 편집기 왼쪽 **⚙️ 프로젝트 설정 → 스크립트 속성 → 속성 추가**

| 속성 | 값 |
|---|---|
| `TEACHER_PW` | `1111` |

코드에 직접 적지 않고 여기에 두는 이유: GitHub 저장소에 비밀번호가 올라가지 않습니다.

## 4. 다시 배포 — **주소가 바뀌지 않게**

**배포 → 배포 관리 → (기존 배포 오른쪽) ✏️ 연필 → 버전: 새 버전 → 배포**

⚠️ "새 배포"를 누르면 **주소가 바뀌어서 학생 링크가 전부 죽습니다.**
반드시 기존 배포를 **편집**해서 새 버전으로 올리세요.

## 5. 확인

1. 선생님 페이지에서 비밀번호를 넣고 들어가 **저장된 퀴즈**가 잘 뜨는지 확인
   (안 뜨면 `index.html`의 비밀번호와 `TEACHER_PW`가 다른 것)
2. 학생용 링크로 들어가 퀴즈가 정상적으로 풀리고 제출되는지 확인

---

## 같이 해두면 좋은 것

- **GitHub 저장소를 비공개(Private)로 바꾸기.**
  지금은 공개라서 누구나 소스에서 Apps Script 주소와 비밀번호 해시를 볼 수 있습니다.
  저장소 → Settings → 맨 아래 Danger Zone → *Change repository visibility* → Private.
  Vercel 무료 요금제도 비공개 저장소를 그대로 배포합니다.
- **비밀번호는 학생이 어깨너머로 봐도 못 외울 정도**로. 세 곳을 같이 바꿔야 합니다:
  `index.html`의 `TEACHER_PW_HASH`(잠금 화면의 "비밀번호 바꾸기"가 값을 만들어 줍니다),
  Apps Script의 `TEACHER_PW`, 그리고 본인 기억.
