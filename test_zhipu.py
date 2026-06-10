import os
from dotenv import load_dotenv, find_dotenv
load_dotenv(find_dotenv(usecwd=True))

from zhipuai import ZhipuAI
key = os.environ.get("ZAI_API_KEY", "")
client = ZhipuAI(api_key=key)

models = [
    "glm-4-flash", "glm-4-flash-250414", "glm-4-flashx",
    "glm-4-air", "glm-4-air-0111", "glm-4-airx",
    "glm-4-plus", "glm-4-plus-0111", "glm-4-0520",
    "glm-4-long", "glm-4-longcontext",
    "glm-4v", "glm-4v-flash", "glm-4v-plus",
    "glm-3-turbo",
    "glm-z1-air", "glm-z1-airx", "glm-z1-flash",
    "characterglm",
    "codegeex-4",
]
for model in models:
    try:
        resp = client.chat.completions.create(
            model=model,
            messages=[{"role": "user", "content": "ok"}],
            max_tokens=3,
        )
        print(f"  OK  {model}: {resp.choices[0].message.content[:20]}")
    except Exception as e:
        code = "?"
        msg = str(e)[:60]
        if "400" in msg:
            code = "model-not-found"
        elif "429" in msg:
            code = "rate-limit"
        elif "401" in msg:
            code = "unauthorized"
        print(f"  {code:18s} {model}")
