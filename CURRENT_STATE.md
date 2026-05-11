# 종소세 2026 현재 상태
> 마지막 업데이트: 2026-05-11

---

## 인프라 구조 (데스크탑 = 개발, 맥미니 = 실행 센터)
```
데스크탑 개발 → git push (jongsotax) → 맥미니 git pull → 크론/봇 실행
```
| 역할 | 위치 |
|------|------|
| 코드 개발·테스트 | Windows 데스크탑 F:\종소세2026 |
| 24시간 실행 센터 | Mac Mini ~/종소세2026 |
| NAS 저장소 | Z:\종소세2026\고객\ (Mac Mini: /Volumes/장성환 또는 /Volumes/장성환-1) |

---

## 안내문 수집 Track 구분

### Track A — 수임동의 완료 고객 (세무대리인 계정으로 일괄 조회)
- 스크립트: `step4_full.py`
- 세무사 사무소 인증서 계정으로 홈택스 일괄 다운로드
- ⚠️ 계정잠금 위험 — 자동 실행 금지, 수동만

### Track B — 신규 고객 (고객 개인 ID/PW로 직접 로그인)
- 스크립트: `step5_individual.py`
- 구글시트 접수명단 col30(홈택스아이디) / col31(홈택스비번) 읽기
- PDF없는 고객만 자동 필터링
- 실행 전 `restart_edge_cdp.bat` 으로 Edge CDP(localhost:9222) 먼저 시작

**Track B 실행 순서:**
```
1. restart_edge_cdp.bat 더블클릭 (Edge CDP 시작)
2. cd F:\종소세2026
3. python step5_individual.py
   → 구글시트에서 홈택스ID/PW 있고 PDF없는 고객 자동 추출
   → Edge로 홈택스 로그인 → 안내문 PDF 다운로드 → parse_anneam() → 구글시트 업데이트
```

### 1명 단독 처리
- `python step6_one_shot.py --name 홍길동 --jumin 000000-0000000 --id ht_id --pw ht_pw`
- 또는 `_run_one.py` (PDF없을 때 수동)

---

## 스크립트 역할 정리

| 파일 | 역할 | 상태 |
|------|------|------|
| `step5_individual.py` | Track B batch: 구글시트 ID/PW → Edge → 안내문 PDF 다운+파싱 | ✅ 작동 |
| `step6_one_shot.py` | Track B 1명: CLI --name --jumin --id --pw | ✅ 작동 |
| `step4_full.py` | Track A batch: 세무대리인 계정 일괄 조회 | ⚠️ 계정잠금 위험 - 수동만 |
| `parse_and_sync_신규.py` | PDF → parse_anneam → 파싱결과.xlsx + 구글시트 | ✅ 작동 |
| `tax_cross_verify.py` | 교차검증 HTML 보고서 생성 (검증보고서_*.html) | ✅ 작동 |
| `print_package.py` | 출력패키지 PDF 생성 (검증보고서+작업결과+안내문+신고서) | ✅ 작동 |
| `airtable_sync_mac.py` | 에어테이블→구글시트 1분 자동싱크 (맥미니 크론) | ✅ 작동 |
| `jongsotaxbot.py` | 텔레그램 봇 — 직원용 작업 자동화 **(Mac Mini LaunchAgent)** | ✅ 작동 |
| `kakao_bank_monitor.py` | 카카오뱅크 입금 감지 → 구글시트 입금체크 **(Windows 전용)** | ✅ 작동 |

---

## 안내문 수집 현황 (2026-05-11 기준)

### ❌ PDF없음 — 처리 필요 (ID/PW 있음, 12명)
| 이름 | 비고 |
|------|------|
| 정도민 | |
| 마금현 | |
| 김진곤 | |
| 지성호 | |
| 배성섭 | |
| 김혜수 | |
| 박성권(KYS) | |
| 황예리(KYS) | |
| 김인덕 | |
| 박수춘 | |
| 장기요 | 신규 |
| 김경오 | 신규 |

→ `python step5_individual.py` 실행하면 위 12명 자동 처리

### ✅ PDF있음: 210명 (Track A/B 완료)

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
- caller에서 stdout 이중 래핑 → import 충돌
- Track A(세무대리인) 자동 실행 → 계정잠금 위험

---

## 시즌 종료 처리 (2026-06-01)
- [x] `airtable_sync_mac.py` — 6/1 자동 중단 내장
- [x] `incometaxbot.py` — 6/1 자동 중단 내장 (SEASON_END)
- [ ] `jongsotaxbot.py` — 6/1 종료 추가 필요

---

## 구글시트
- ID: `1oh31k00Oa2lZWvu5fnBRVmurdlll1YEG8Fefi5FRfBI`
- 접수명단 시트 (gid=245653883): col30=홈택스아이디, col31=홈택스비번
- 에어테이블 1분 싱크 → Mac Mini 크론 (`~/airtable_sync.log`)

## NAS
- IP: `192.168.0.100` (DS920+)
- SMB 계정: `admin` (changmini 아님)
- Mac Mini 마운트: `/Volumes/장성환`, `/Volumes/장성환-1`
- HDD 절전: **없음** (2026-05-11 비활성화 — 봇 SMB 연결 안정화)
