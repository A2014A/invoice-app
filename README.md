# מערכת חשבוניות — יהודה קורץ

אפליקציית ווב להפקת חשבוניות, קבלות וחשבוניות+קבלה.

## פריסה על Render.com

1. העלה את התיקייה ל-GitHub (repository חדש)
2. היכנס ל-render.com
3. לחץ "New Web Service"
4. חבר את ה-GitHub repository
5. Render יזהה אוטומטית את render.yaml
6. לחץ "Deploy"

## הרצה מקומית

pip install -r requirements.txt
python app.py

## קבצים

- app.py — שרת Flask + הפקת מסמכי Word
- templates/index.html — ממשק משתמש
- requirements.txt — תלויות
- render.yaml — הגדרות פריסה
- clients.json — רשימת לקוחות (נוצר אוטומטית)
