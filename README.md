# 여깁니다유 랜덤박스/라방 운영 프로그램 v94

## 실행
```bash
npm install
npm run dev
```

## 배포
GitHub에 전체 폴더를 올린 뒤 Vercel에서 연결하세요.

## 환경변수
Vercel Project Settings > Environment Variables에 추가:
```env
GEMINI_API_KEY=제미나이_API_KEY
GEMINI_MODEL=gemini-2.5-flash
VITE_SUPABASE_URL=Supabase_Project_URL
VITE_SUPABASE_ANON_KEY=Supabase_Anon_Key
```

## Supabase
배포 전 Supabase SQL Editor에서 `supabase_setup.sql` 전체를 실행하세요.

## v94 변경사항
- 랜덤스쿱 탭 제거
- 라방주문 상태 흐름 정리: 미입금 → 입금확인 → 정산후킵/송장입력 → 출고완료
- 라방상품 추가 시 본재고에서 라방재고로 수량 이동
- 주문 저장 시 미입금 상태로 라방재고 예약
- 입금확인 상태부터 라방 대시보드 매출/순수익 반영
- 정산후킵 D-Day 표시 및 D-2/출고필요 강조
- 회원 통합관리: 고객명+전화번호 뒷4자리 기준
- 같은 회원 미출고 주문 합배송 묶기
- 주문자 정보: 우편번호/기본주소/상세주소/전화번호
- 택배접수 엑셀: 박스무게 2/5, 박스부피 60/80/100, 생활용품, 메모 반영
- 라방 정산서 PDF/엑셀 버튼 유지

## v100 변경사항
- 라방 구매자 주문 생성 영역에서 회원 검색/불러오기/회원저장 가능
- 주문 생성 시 포인트 입력 및 전액 사용 기능 추가
- 포인트 사용 시 주문 정산서 PDF/엑셀에 포인트 차감 반영
- 회원 보유 포인트와 사용 누적 포인트를 Supabase에 저장
- 주문관리 상태 변경은 주문 행별 드롭다운에서 처리하도록 UI 정리
- 주문관리의 혼란스러운 상태 필터/일괄 입금확인 버튼 제거
- 재고 1개 제외 체크 시 검색 입력값이 사라지는 문제 수정

## Supabase 추가 실행 SQL
이미 기존 SQL을 실행했더라도 아래 컬럼 추가가 필요합니다.
```sql
alter table live_members add column if not exists used_points integer default 0;
alter table live_orders add column if not exists used_points integer default 0;
```
