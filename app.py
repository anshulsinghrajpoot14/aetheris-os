# ============================================================
# 4. DIRECT REASONING ENGINE (LIVE DIAGNOSTIC MODE)
# ============================================================
def ask_aetheris(system_instruction: str, user_prompt: str, max_tokens=2500) -> str:
    log_errors = []

    # 1. Groq Inference
    if GROQ_API_KEY and Groq:
        try:
            client = Groq(api_key=GROQ_API_KEY, timeout=20.0)
            for model_name in ["llama-3.3-70b-versatile", "llama-3.1-8b-instant"]:
                try:
                    res = client.chat.completions.create(
                        model=model_name,
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
                except Exception as e_inner:
                    log_errors.append(f"Groq [{model_name}]: {str(e_inner)}")
        except Exception as e_outer:
            log_errors.append(f"Groq Client Init: {str(e_outer)}")
    else:
        log_errors.append("Groq: API Key missing or library not loaded")

    # 2. Gemini REST Fallback (Direct v1beta endpoint)
    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            clean_key = GEMINI_API_KEY.replace('"', '').replace("'", "").strip()
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={clean_key}"
            headers = {"Content-Type": "application/json"}
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": f"{system_instruction}\n\nUser Question: {user_prompt}"}
                        ]
                    }
                ],
                "generationConfig": {
                    "temperature": 0.3,
                    "maxOutputTokens": max_tokens
                }
            }
            resp = requests.post(url, headers=headers, json=payload, timeout=20)
            if resp.status_code == 200:
                data = resp.json()
                candidates = data.get("candidates", [])
                if candidates:
                    parts = candidates[0].get("content", {}).get("parts", [])
                    if parts and "text" in parts[0]:
                        return parts[0]["text"].strip()
            else:
                log_errors.append(f"Gemini HTTP {resp.status_code}: {resp.text[:120]}")
        except Exception as e_gemini:
            log_errors.append(f"Gemini Connection: {str(e_gemini)}")
    else:
        log_errors.append("Gemini: API Key missing or requests library unavailable")

    return "⚠️ Diagnostic Breakdown:\n" + "\n".join(f"• {err}" for err in log_errors)
