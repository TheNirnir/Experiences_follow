# מעקב התנסויות צוות

אפליקציית ווב פשוטה למעקב אחרי התנסויות הדרכתיות ופיקודיות לפי שלבי מעגל החניכה.

- **האתר** (`index.html`, `config.js`) הוא דף סטטי. אפשר להעלות אותו ל-GitHub Pages.
- **הנתונים** נשמרים בגיליון Google, דרך API קטן ב-Apps Script (`Code.gs`).

כל עוד `API_URL` ב-`config.js` ריק, האתר רץ **במצב הדגמה**: הנתונים נשמרים רק בדפדפן הנוכחי.

## 1. חיבור לגיליון Google (פעם אחת)

1. פתחו גיליון חדש: [sheets.new](https://sheets.new).
2. בתפריט: **תוספים ← Apps Script**.
3. החליפו את תוכן `Code.gs` בעורך בתוכן הקובץ `Code.gs` מכאן, ושמרו.
4. בחרו את הפונקציה `setup` ולחצו **Run**. אשרו את ההרשאות (Advanced ← Go to … ← Allow).
   בגיליון ייווצרו הלשוניות **אנשים**, **שלבים**, **התנסויות**.
5. **Deploy ← New deployment**, סוג **Web app**:
   - Execute as: **Me**
   - Who has access: **Anyone**
6. העתיקו את הקישור (נגמר ב-`/exec`) והדביקו אותו ב-`config.js`:
   ```js
   window.CONFIG = { API_URL: 'https://script.google.com/macros/s/XXXX/exec' };
   ```

## 2. העלאה ל-GitHub Pages

```bash
git init
git add index.html config.js Code.gs appsscript.json README.md
git commit -m "Experience tracker"
git branch -M main
git remote add origin https://github.com/<user>/<repo>.git
git push -u origin main
```

ב-GitHub: **Settings ← Pages ← Source: Deploy from a branch ← main / root ← Save**.
אחרי כדקה האתר יהיה זמין בכתובת `https://<user>.github.io/<repo>/`. בטלפון אפשר לבחור "הוספה למסך הבית".

> כל מי שיש לו את הקישור יכול לצפות ולערוך.

## עריכת שמות, שלבים והודעות

הכול נערך ישירות בגיליון:

- **אנשים**: שם אחד בכל שורה.
- **שלבים**: `סוג | מספר | שם שלב | הודעה`. הסוג הוא `הדרכתית` או `פיקודית`, ולכל סוג יכול להיות מספר שלבים אחר.
  בהודעה אפשר לכתוב `{שם}` ו-`{התנסות}`, והם יוחלפו אוטומטית בשם החניך ובתיאור ההתנסות.
- **התנסויות**: נכתבת על ידי האפליקציה.

## עדכון קוד ה-API

אחרי שינוי ב-`Code.gs`: **Deploy ← Manage deployments ← ✏️ ← Version: New version ← Deploy**. הקישור נשאר אותו קישור.
