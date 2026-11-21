# 🚀 Morning Arena

אפליקציית התארגנות בוקר לאלון ואיתן — בנויה על Google Apps Script + Google Sheets.

## מה זה?

לוח בקרה גיימיפיקציה לילדים שעוזר להם להתארגן בבוקר עצמאית.  
כל ילד בוחר פרופיל, מסמן משימות, וצובר XP — עם אנימציות, חגיגות וסטריק יומי.

## פיצ'רים

- 👤 **פרופילים** — אלון (ציאן/סגול) ואיתן (ירוק/כתום)
- ✅ **10 משימות בוקר** עם checkmark מונפש ו-haptic feedback
- 🔒 **נעילת 5 דקות** — לא אפשר לבטל משימה אחרי 5 דקות
- 🏆 **XP + מעגל התקדמות** — גיימיפיקציה מלאה
- 🔥 **סטריק יומי** — ימים רצופים של השלמת משימות
- 📊 **היסטוריה** — ימים / שבוע / חודש
- 🎉 **חגיגת סיום** עם קונפטי ואפקטי קול
- 🌌 **עיצוב** — dark space theme, גבולות גרדיאנט מונפשים, חלקיקים

## טכנולוגיות

- **Frontend:** HTML/CSS/JS (Google Apps Script HtmlService)
- **Backend:** Google Apps Script
- **Database:** Google Sheets (`Logs_V2`)
- **Libraries:** canvas-confetti

## התקנה

1. פתח [Google Apps Script](https://script.google.com) וצור פרויקט חדש
2. העלה את `Code.gs` ו-`Index.html`
3. קשר ל-Google Sheet עם גיליון בשם `Logs_V2`
4. פרוס כ-Web App (Execute as: Me, Who has access: Anyone)

## קבצים

| קובץ | תיאור |
|------|--------|
| `Code.gs` | Backend — קריאת/כתיבת Sheets, לוגיקת נעילה |
| `Index.html` | Frontend — כל ה-UI, אנימציות, לוגיקת לקוח |
