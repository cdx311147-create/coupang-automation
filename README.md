# 🛒 E-commerce Automation Tools
> 이커머스 MD 업무 자동화 시스템 모음 (Python / Selenium / VBA)

## 📌 프로젝트 개요
패션 이커머스 MD로 근무하며 수천 개 SKU의 가격·재고·발주 데이터를
수기로 관리하는 비효율을 해결하기 위해 직접 개발한 자동화 시스템입니다.

---

## 🔧 프로젝트 목록

### 1. 쿠팡 Ads Center 실시간 가격 수집기
`coupang_ASP.py`
- **배경:** 쿠팡 로켓배송 판매가는 공식 API(파트너스/WING/서플라이어허브)에서 조회 불가
- **해결:** Ads Center 수동 상품 설정 화면을 Selenium으로 자동화하여 전 SKU 실시간 가격 수집
- **기술:** Python, Selenium, openpyxl
- **성과:** SKU 8,000개 이상 실시간 가격 수집 자동화, 데이터 누락률 0% 달성

### 2. 네이버 쇼핑 최저가 소싱 자동화
`지식쇼핑 API.py`
- **배경:** 수백 개 상품의 최저가를 수기로 확인하는 반복 업무 개선
- **해결:** 네이버 Open API + Selenium 2단계 구조로 해외·리셀러 필터링 후 실질 최저가 수집
- **기술:** Python, Selenium, Naver Open API, openpyxl
- **성과:** 소싱 데이터 수집 시간 90% 단축

---

## 🛠 기술 스택
| 분류 | 기술 |
|------|------|
| Language | Python, VBA |
| Automation | Selenium, openpyxl |
| API | Coupang Partners API, Naver Open API |
| Tools | Chrome WebDriver |

---

## ⚠️ 주의사항
- 본 코드는 **본인 계정의 본인 데이터**를 수집하기 위한 용도로 개발되었습니다.

---

## 👤 개발자
- **소속:** (주)스타인터내셔널 MD 최은태
- **개발 목적:** 이커머스 현업 업무 자동화
