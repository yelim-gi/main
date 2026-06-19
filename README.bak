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
