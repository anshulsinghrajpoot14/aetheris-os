def ask_aetheris(system_instruction: str, user_prompt: str, max_tokens=2500) -> str:
    errors = []

    # 1. Groq Test
    if GROQ_API_KEY and Groq:
        try:
            client = Groq(api_key=GROQ_API_KEY, timeout=20.0)
            res = client.chat.completions.create(
                model="llama-3.1-8b-instant",
                messages=[
                    {"role": "system", "content": system_instruction},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=0.3,
                max_tokens=max_tokens
            )
            if res.choices and res.choices[0].message.content:
                return res.choices[0].message.content.strip()
        except Exception as e:
            errors.append(f"Groq Error: {str(e)}")
    else:
        errors.append("Groq Error: GROQ_API_KEY missing or invalid")

    # 2. Gemini REST Test
    if GEMINI_API_KEY and REQUESTS_OK:
        try:
            clean_key = GEMINI_API_KEY.replace('"', '').replace("'", "").strip()
            url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={clean_key}"
            payload = {
                "contents": [{"parts": [{"text": f"{system_instruction}\n\nUser Question: {user_prompt}"}]}],
                "generationConfig": {"temperature": 0.3, "maxOutputTokens": max_tokens}
            }
            resp = requests.post(url, json=payload, timeout=20)
            if resp.status_code == 200:
                txt = resp.json()["candidates"][0]["content"]["parts"][0]["text"].strip()
                if txt:
                    return txt
            else:
                errors.append(f"Gemini HTTP {resp.status_code}: {resp.text[:120]}")
        except Exception as e:
            errors.append(f"Gemini Error: {str(e)}")
    else:
        errors.append("Gemini Error: GEMINI_API_KEY missing or invalid")

    return f"⚠️ Diagnostics:\n" + "\n".join(errors)
