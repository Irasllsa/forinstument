import os, requests
from PIL import Image
from io import BytesIO

os.makedirs("assets/img", exist_ok=True)

# Curated Pexels photos (free to use) relevant to research / education / academia
images = {
    "cover":        "https://images.pexels.com/photos/267885/pexels-photo-267885.jpeg?auto=compress&cs=tinysrgb&w=1400",   # graduation/university
    "students":     "https://images.pexels.com/photos/3184465/pexels-photo-3184465.jpeg?auto=compress&cs=tinysrgb&w=1200",# team discussion
    "library":      "https://images.pexels.com/photos/256541/pexels-photo-256541.jpeg?auto=compress&cs=tinysrgb&w=1200",  # library shelves
    "books":        "https://images.pexels.com/photos/159711/books-bookstore-book-reading-159711.jpeg?auto=compress&cs=tinysrgb&w=1200",
    "writing":      "https://images.pexels.com/photos/590016/pexels-photo-590016.jpeg?auto=compress&cs=tinysrgb&w=1200",   # laptop/notebook
    "study":        "https://images.pexels.com/photos/8617843/pexels-photo-8617843.jpeg?auto=compress&cs=tinysrgb&w=1200",# student studying
    "presentation": "https://images.pexels.com/photos/3184292/pexels-photo-3184292.jpeg?auto=compress&cs=tinysrgb&w=1200",# meeting/presenting
    "data":         "https://images.pexels.com/photos/590022/pexels-photo-590022.jpeg?auto=compress&cs=tinysrgb&w=1200",  # charts/analytics
    "idea":         "https://images.pexels.com/photos/355948/pexels-photo-355948.jpeg?auto=compress&cs=tinysrgb&w=1200",  # lightbulb
    "magnify":      "https://images.pexels.com/photos/6238297/pexels-photo-6238297.jpeg?auto=compress&cs=tinysrgb&w=1200",# analysis
    "teacher":      "https://images.pexels.com/photos/5212345/pexels-photo-5212345.jpeg?auto=compress&cs=tinysrgb&w=1200",# classroom
    "thinking":     "https://images.pexels.com/photos/4144923/pexels-photo-4144923.jpeg?auto=compress&cs=tinysrgb&w=1200",# thinking/notes
    "success":      "https://images.pexels.com/photos/3771074/pexels-photo-3771074.jpeg?auto=compress&cs=tinysrgb&w=1200",# success/celebration
    "campus":       "https://images.pexels.com/photos/207692/pexels-photo-207692.jpeg?auto=compress&cs=tinysrgb&w=1200",  # writing close
}

headers = {"User-Agent": "Mozilla/5.0"}
ok, fail = [], []
for name, url in images.items():
    try:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        img = Image.open(BytesIO(r.content)).convert("RGB")
        path = f"assets/img/{name}.jpg"
        img.save(path, "JPEG", quality=88)
        ok.append((name, img.size))
    except Exception as e:
        fail.append((name, str(e)[:60]))

print("OK:")
for n, s in ok: print(f"  {n}: {s}")
print("FAIL:")
for n, e in fail: print(f"  {n}: {e}")
