# 종소세 2026 현재 상태
> 마지막 업데이트: 2026-05-02 (incometaxbot Windows 전환)

---

## 인프라 구조 (데스크탑 = 개발, 맥미니 = 실행 센터)
```
데스크탑 개발 → git push (jongsotax) → 맥미니 git pull → 크론/봇 실행
```
| 역할 | 위치 |
|------|------|
| 코드 개발·테스트 | Windows 데스크탑 F:\종소세2026 |
| 24시간 실행 센터 | Mac Mini ~/종소세2026 |
| NAS 저장소 | Z:\종소세2026\고객\ |

## Git Repo 구조
| Repo | 용도 |
|------|------|
| changsh777/jongsotax | 종소세 자동화 전체 (데스크탑+맥미니 공유) |
| changsh777/taxbot-automation | 부가세·원천세 봇 (~/taxbot/) |
| changsh777/macmini-bots | 맥미니 텔레그램 봇·크론 스크립트 모음 |

---

## 스크립트 역할 정리

| 파일 | 역할 | 상태 |
|------|------|------|
| `신규고객처리.py` | 고객구분=신규: 고객 ID/PW 로그인 → PDF 다운 | ✅ 작동 |
| `기존고객처리.py` | 고객구분=기존: 세무사 인증서 로그인 → PDF 다운 | ⚠️ 계정잠금 위험 - 수동만 |
| `parse_and_sync_신규.py` | PDF → parse_anneam → 파싱결과.xlsx + 구글시트 | ✅ 작동 |
| `airtable_sync_mac.py` | 에어테이블→구글시트 1분 자동싱크 (맥미니 크론) | ✅ 작동 |
| `show_status.py` | 전체 고객 처리 현황 출력 | ✅ 작동 |
| `incometaxbot.py` | Telegram봇 @incometax777_bot — 신규접수 감지→다운+파싱 자동화 (Windows 실행) | ✅ 완성 |
| `_run_one.py` | incometaxbot 서브프로세스용 — Edge CDP 다운+파싱 1명 처리 | ✅ 완성 |

---

## 신규 고객 처리 현황 (2026-05-02 기준)

### ✅ 완료 (PDF + 파싱 + 구글시트)
| 이름 | 수입 | 비고 |
|------|------|------|
| 신정숙 | 34,814,000 | 간편장부 / 기준경비율 |
| 한효성 | 42,413,744 | 간편장부 / 기준경비율 (기존고객이지만 ID/PW로 처리) |

### ❌ 미처리
| 이름 | 메모 |
|------|------|
| 정도민 | 홈택스 아이디 있음 |
| 김태윤 | 홈택스 아이디 있음 |
| 지성호 | 홈택스 아이디 있음 |
| 김경필 | 홈택스 아이디 있음 |
| 이재윤 | 홈택스 아이디 있음 |
| 유영주 | 홈택스 자료 없음 (소득 없을 가능성) |

---

## 기존 고객 처리 현황

### PDF 있는 것
- 김병수, 채민희, 이민수, 이명회, 한효성

### 수입 없음 → PDF 미처리
- 마금현, 김진곤, 이윤경, 양태석, 장은향, 한두열, 진현오, 고홍, REEVE SAMANTHA EMMA
- 장성환(본인), 박수경 = 홈택스 아이디 없음

---

## 홈택스 로그인 핵심 패턴 (삽질 방지)

```python
# 로그아웃
logout_hometax(page)  # HTML 모달 → [id*=_btn_confirm].click()

# 로그인 (신규/아이디 방식)
login_hometax_id(page, hometax_id, hometax_pw, jumin_raw="XXXXXX-XXXXXXX")
# - ID: #mf_txppWframe_loginboxFrame_iptUserId
# - 2차인증: [id*=UTXPPABC12][id*=iptUserJuminNo1/2], 확인: [id*=UTXPPABC12][id*=trigger46]

# 신고도움서비스 이동 (goto 금지! SPA)
page.evaluate("document.getElementById('menuAtag_4103080000').onclick()")
```

### ❌ 금지
- `page.goto(tmIdx URL)` → "로그인 정보가 없습니다"
- caller에서 stdout 이중 래핑 → 신규고객처리.py import 충돌
- 기존고객처리.py 자동 실행 → 계정잠금 위험 (어제 발생)

---

## incometaxbot 실행 방법 (Windows 데스크탑)

```
# 터미널에서:
cd F:\종소세2026
python incometaxbot.py

# Edge CDP는 PDF 다운로드 필요할 때만 미리 열어두기:
python launch_edge.py
```

**흐름 요약:**
- n8n → Telegram "@incometax777_bot" 채팅 → "홍길동님 신규 접수되었습니다."
- 봇이 구글시트 조회 → NAS PDF 확인
- PDF 있음 → parse_and_sync_신규.py 직접 실행 (Edge 불필요)
- PDF 없음 → _run_one.py 서브프로세스 → Edge CDP 로그인+다운+파싱

---

## 시즌 종료 처리 (2026-06-01)
- [x] `airtable_sync_mac.py` — 6/1 자동 중단 내장
- [x] `incometaxbot.py` — 6/1 자동 중단 내장 (SEASON_END)
- [ ] `jongsotaxbot.py` — 6/1 종료 추가 필요

---

## 다음 할 일
1. n8n 워크플로우 — 에어테이블 신규접수 시 @incometax777_bot 채팅으로 Telegram 메시지 전송 추가
2. 신규 미처리 고객 건바이건 처리 (정도민·김태윤·지성호·김경필·이재윤)
3. 기존 고객 PDF — 세무사 수동 로그인 방식으로 조심히
4. jongsotaxbot.py 6/1 종료 처리

---

## 구글시트
- ID: `1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI`
- 시트: `접수명단`
- 에어테이블 1분 싱크 → Mac Mini 크론 (`~/airtable_sync.log` 로그)
