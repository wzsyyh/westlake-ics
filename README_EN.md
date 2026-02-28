# westlake-ics

Convert your Westlake University SIS course schedule into calendar formats used by popular apps.

| Script | Output | Works with |
|--------|--------|------------|
| `ps2ics.py` | `.ics` | Apple Calendar, Google Calendar, Outlook |
| `ps2wakeup.py` | `.csv` | WakeUp Schedule App |

**No third-party libraries required. Python 3.6+ only.**

If this helps you, please consider giving it a ⭐ Star so more students can find it!

---

## How to use

### Step 1: Get the project

**Option A: Download ZIP (recommended if you're not familiar with Git)**

Click the green **Code** button on the top right of this page → **Download ZIP**, then unzip it.

**Option B: git clone**

```bash
git clone https://github.com/wzsyyh/westlake-ics.git
cd westlake-ics
```

---

### Step 2: Download your schedule from SIS

1. Open the [SIS course system](https://sis.westlake.edu.cn/psc/CSPRD_3/EMPLOYEE/SA/c/SSR_STUDENT_FL.SSR_MD_SP_FL.GBL?Action=U&MD=Y&GMenu=SSR_STUDENT_FL&GComp=SSR_START_PAGE_FL&GPage=SSR_START_PAGE_FL&scname=CS_SSR_MANAGE_CLASSES_NAV) (login required)
2. Click **查看我的课表** (View My Schedule) in the left sidebar
3. Select the **按课程** (By Course) tab at the top
4. Click the **download button** in the top-right corner (see the red arrow in the screenshot below)

![Download button location](ref.png)

5. Save the file — it defaults to `ps.xls`. Move it into this project folder.

---

### Step 3: Run the script

Open a terminal, navigate to the project folder, and run:

**For Apple Calendar / Google Calendar / Outlook:**

```bash
python3 ps2ics.py ps.xls
```

This generates `schedule.ics`.

**For WakeUp Schedule App:**

```bash
python3 ps2wakeup.py ps.xls
```

This generates `wakeup.csv`.

You can also specify a custom output path with `-o`:

```bash
python3 ps2ics.py ps.xls -o ~/Desktop/schedule.ics
```

---

### Step 4: Import into your calendar

#### Apple Calendar

Double-click `schedule.ics`. A confirmation dialog will appear — click **Import**.

#### Google Calendar

1. Open [Google Calendar](https://calendar.google.com) in a browser
2. Click the gear icon (top right) → **Settings** → **Import & export** → **Import**
3. Select `schedule.ics` and click **Import**

#### Outlook

Double-click `schedule.ics` and confirm. Or manually: **File → Open & Export → Import/Export → Import an iCalendar (.ics)**

#### WakeUp Schedule App

WakeUp's CSV format does not carry period time information, so you need to **configure the period schedule once** before importing. You won't need to repeat this every semester.

**① Set up period times (one-time setup)**

In WakeUp, go to **Schedule Settings → Period Times**, set the number of periods to 13, and fill in the table below:

| Period | Start | End | Period | Start | End |
|--------|-------|-----|--------|-------|-----|
| 1 | 08:00 | 08:45 | 8 | 15:10 | 15:55 |
| 2 | 08:50 | 09:35 | 9 | 16:10 | 16:55 |
| 3 | 09:50 | 10:35 | 10 | 17:00 | 17:45 |
| 4 | 10:40 | 11:25 | 11 | 18:30 | 19:15 |
| 5 | 11:30 | 12:15 | 12 | 19:20 | 20:05 |
| 6 | 13:30 | 14:15 | 13 | 20:10 | 20:55 |
| 7 | 14:20 | 15:05 | | | |

**② Import courses**

In WakeUp: top-right menu → **Import Schedule** → **From CSV file** → select `wakeup.csv`

---

## Notes

- **EST2005 (Data Structures & Algorithm Design in C++)** has two duplicate Friday entries in the raw SIS data (rooms H6-303 and H6-305). You can manually delete one after importing.
- All times are written in **Beijing Time** (Asia/Shanghai), so there will be no timezone offset issues across platforms.
- If any course times appear incorrect, make sure the `ps.xls` you downloaded is for the current semester.
