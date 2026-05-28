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
