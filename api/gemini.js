
export default async function handler(req, res) {
  try {
    if (req.method !== "POST") {
      return res.status(405).json({ error: "POST 요청만 사용할 수 있어요." });
    }

    const apiKey = String(process.env.GEMINI_API_KEY || process.env.VITE_GEMINI_API_KEY || "").trim();
    if (!apiKey) {
      return res.status(500).json({
        error: "GEMINI_API_KEY 환경변수가 비어 있어요. Vercel Project Settings > Environment Variables에 GEMINI_API_KEY를 추가해주세요.",
      });
    }

    const { message, context, imageBase64, mimeType, task } = req.body || {};
    if (!message || !String(message).trim()) {
      return res.status(400).json({ error: "메시지가 비어 있어요." });
    }

    const systemPrompt = task === "inventory_image_parse" || task === "inventory_text_parse"
      ? `너는 캐릭터 굿즈 쇼핑몰 입고자료 분석 비서다. 거래명세서, 영수증, 주문내역서, 상품 사진에서 재고 등록에 필요한 정보를 추출한다. 반드시 JSON 배열만 반환한다. 마크다운, 설명문, 주석은 쓰지 않는다. 필드는 name, char1, char2, category, stock, wholesale, retail 만 사용한다. 모르는 값은 빈 문자열 또는 0으로 둔다. 여러 상품/여러 장의 자료가 있으면 전부 추출한다.`
      : `너는 캐릭터 랜덤박스 쇼핑몰 '여깁니다유' 운영 비서다. 고객 요청사항 분석, 랜덤박스/랜덤스쿱 조합 추천, 원하는 캐릭터 중심 상품 추천, 재고 부족 체크, 상품명/카테고리/수량 수정 제안, 택배접수 정리, 재주문 중복 방지, 일정관리 조언을 한다. 실제 재고 상품명을 최대한 그대로 사용하고, 수동박스로 내보내기 쉽게 추천 상품명은 줄마다 하나씩 써라. 재고 0개 상품은 추천하지 말고, 목표 마진율과 판매가를 고려해라. 데이터 변경은 사용자 확인 후에만 반영된다. 답변은 한국어로 간단하고 구체적으로 한다.`;

    const finalPrompt = `${systemPrompt}\n\n현재 프로그램 데이터:\n${JSON.stringify(context || {}, null, 2)}\n\n사용자 요청:\n${message}`;

    const models = [
      process.env.GEMINI_MODEL,
      "gemini-2.5-flash",
      "gemini-2.0-flash",
      "gemini-1.5-flash",
    ].filter(Boolean);

    let lastData = null;
    for (const model of models) {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(apiKey)}`;
      const parts = [{ text: finalPrompt }];
      if (imageBase64) {
        parts.push({
          inline_data: {
            mime_type: mimeType || "image/png",
            data: String(imageBase64).includes(",") ? String(imageBase64).split(",").pop() : imageBase64,
          },
        });
      }

      const upstream = await fetch(url, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ contents: [{ role: "user", parts }] }),
      });

      lastData = await upstream.json().catch(() => ({}));
      if (upstream.ok) {
        const text = lastData?.candidates?.[0]?.content?.parts?.map((p) => p.text).join("\n") || "";
        return res.status(200).json({ text, model });
      }
    }

    return res.status(500).json({ error: "Gemini 호출 실패", detail: lastData });
  } catch (error) {
    return res.status(500).json({ error: error?.message || String(error) });
  }
}
