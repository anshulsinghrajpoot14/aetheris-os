# ============================================================
# 4. DIRECT REASONING ENGINE (AUTO-DISCOVERY & BULLETPROOF FALLBACK)
# ============================================================
def ask_aetheris(system_instruction: str, user_prompt: str, max_tokens=2500) -> str:
    # 1. Try Groq with Auto-Discovery of available models
    if GROQ_API_KEY and Groq:
        try:
            client = Groq(api_key=GROQ_API_KEY, timeout=25.0)
            
            # Step A: Check which models are actually available on your key
            available_models = []
            try:
                m_list = client.models.list()
                available_models = [m.id for m in m_list.data if "whisper" not in m.id.lower() and "guard" not in m.id.lower()]
            except Exception:
                pass
            
            # Fallback priority list
            candidate_models = available_models + ["llama3-8b-8192", "llama3-70b-8192", "mixtral-8x7b-32768", "gemma2-9b-it"]
            
            for m_id in candidate_models:
                try:
                    res = client.chat.completions.create(
                        model=m_id,
                        messages=[
                            {"role": "system", "content": system_instruction},
                            {"role": "user", "content": user_prompt}
                        ],
                        temperature=0.3,
                        max_tokens=max_tokens
                    )
                    if res.choices and res.choices[0].message.content:
                        txt = res.choices[0].message.content.strip()
                        if len(txt) > 0:
                            return txt
                except Exception:
                    continue
        except Exception:
            pass

    # 2. Try Gemini Stable v1 Endpoint
    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            clean_key = GEMINI_API_KEY.replace('"', '').replace("'", "").strip()
            for ep in [
                f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={clean_key}",
                f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key={clean_key}"
            ]:
                try:
                    headers = {"Content-Type": "application/json"}
                    payload = {
                        "contents": [{"parts": [{"text": f"{system_instruction}\n\nUser Question: {user_prompt}"}]}],
                        "generationConfig": {"temperature": 0.3, "maxOutputTokens": max_tokens}
                    }
                    resp = requests.post(ep, headers=headers, json=payload, timeout=20)
                    if resp.status_code == 200:
                        data = resp.json()
                        candidates = data.get("candidates", [])
                        if candidates:
                            parts = candidates[0].get("content", {}).get("parts", [])
                            if parts and "text" in parts[0]:
                                return parts[0]["text"].strip()
                except Exception:
                    continue
        except Exception:
            pass

    return "System ready. Please verify your connection."
