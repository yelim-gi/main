# 여깁니다유 운영 프로그램 v162

## 실행
npm install
npm run dev

## v162 수정
- 입금후킵/입금후합배송 주문 정산서에 킵 자동 발송 날짜 표시
- 킵 날짜는 상태 변경일 기준, 당일 포함 계산
- 라방 설정의 킵기간 변경 시 기존 킵 주문의 만료일도 재계산
- 이벤트경품 상품 선택을 드롭다운 대신 검색/선택 방식으로 변경

## Supabase
supabase_setup.sql을 한 번 실행해주세요.


## v163 중요
이벤트경품/킵 자동발송 기능을 쓰려면 배포 전에 `SUPABASE_먼저실행.sql`을 Supabase SQL Editor에서 한 번 실행하세요.
- event_prizes 테이블 생성
- live_orders 킵 관련 컬럼 추가
- schema cache reload 포함
