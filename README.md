# 여깁니다유 랜덤박스/라방 운영 프로그램 v102

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
이미 실행했다면 v100 포인트 컬럼까지 포함되어 있는지 확인하세요.

## v102 변경사항
- 라방 주문관리에서 주문 수정 기능 추가
- 주문 수정 시 기존 라방 예약수량을 반영해서 재고 꼬임 방지
- 주문 취소/삭제/라방 삭제 시 사용 포인트도 회원 포인트로 복구
- 배송비 적용 체크 시 선결제/포인트 여부와 관계없이 배송비 반영
- 라방 택배접수 엑셀은 선택 라방의 입금확인/송장입력 주문 전체를 한 엑셀로 생성
- 구매자 주문생성/통합회원관리에서 다음 우편번호 검색 추가
- 정산서 엑셀은 `public/invoice_template.xlsx` 템플릿 양식을 기반으로 생성
- 정산서 PDF는 A4 문서형 양식으로 정리하고 포인트 차감 포함
- 회원 포인트 전액 사용/차감/복구 반영
