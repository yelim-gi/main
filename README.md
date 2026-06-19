# 여깁니다유 랜덤박스/라방 운영 프로그램 v98 검증본

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
기존 데이터 삭제 없이 `create table if not exists` / `alter table add column if not exists` 위주로 반영됩니다.

## v98 검증/수정사항
- `npm run build` 성공 확인
- 라방주문은 `live_sessions / live_members / live_orders` 테이블 구조 기준으로 동작
- 존재하지 않는 `live_products / live_order_items / live_keep_items` 테이블 참조 없음 확인
- 라방 삭제/재고복구 버튼 추가
- 라방 주문 삭제 시, 취소 전 주문이면 라방 남은수량 복구 후 삭제
- 주문 취소 시 라방 남은수량 복구
- 라방상품 삭제 시 본재고 복구
- 합배송 묶기: 고객명 + 전화번호 뒷 4자리 기준, 미출고 주문 묶음
- 택배접수 엑셀은 합배송 묶음 기준 1건으로 내보냄
- 라방 정산서 PDF/엑셀 버튼 유지
- 랜덤스쿱 탭/화면 제거 상태 유지

## 주의
실제 Supabase 프로젝트에는 직접 접속 테스트를 할 수 없으므로, 배포 후 라방 생성 → 상품 추가 → 주문 저장 → 입금확인 → 취소/삭제 테스트를 한 번만 해주세요.
