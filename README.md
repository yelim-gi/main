# 랜덤박스 운영 프로그램 Vercel 배포용

## 환경변수

Vercel Project Settings > Environment Variables에 아래 값을 추가하세요.

```env
GEMINI_API_KEY=제미나이_API_KEY
GEMINI_MODEL=gemini-2.5-flash

VITE_SUPABASE_URL=Supabase_Project_URL
VITE_SUPABASE_ANON_KEY=Supabase_Anon_Key
```

## 실행

```bash
npm install
npm run dev
```

## 빌드

```bash
npm run build
```

## 주요 기능

- 대시보드
- 정산매입매출
- AI사입입고분석
- 재고관리
- 수동박스
- 랜덤스쿱
- 주문관리
- 택배접수
- Gemini 운영 비서
- Vercel API Route: `/api/gemini`

## AI사입입고분석

- 이미지/PDF 여러 개 업로드 가능
- 거래명세서/주문목록 여러 건 복붙 가능
- 분석 결과를 표에서 직접 수정 후 재고 반영 가능

## 택배접수

- 엑셀 다운로드 파일명: `택배접수.xlsx`
- 첫 행부터 데이터만 저장
- 다운로드 후 목록 삭제 여부 확인


## v87 수정사항

- 수동박스 입력값 통합
  - 박스수
  - 판매가
  - 수수료율 기본 6.4
  - 목표마진율
  - 고객명
  - 메모
- AI 수동박스 추천은 위 통합 입력값을 기준으로 사용
- 중복 판매가/목표마진율 입력칸 제거
- 주문관리 선택 주문상품 목록을 주문표 아래쪽으로 배치
- 주문접수/출고완료 표 높이를 크게 조정


## v88 수정사항

- 택배접수의 재주문 전화번호 확인 기능 제거
- 재주문확인 컬럼/버튼 제거
- 같은 메뉴를 다시 클릭할 때 불필요한 화면 갱신 방지


## v89 수정사항

- 캐릭터1/캐릭터2 선택창을 페이지 안에 갇히지 않는 진짜 모달로 변경
- 재고 1개 제외 체크 시 상품명 검색어/카테고리/캐릭터 선택값 유지
- 상품 추가 시 왼쪽 상품검색 목록 스크롤 위치 유지
