# westlake-ics

将西湖大学教务系统（SIS）的课程表自动转换为主流日历格式：

| 脚本 | 输出 | 适用 |
|------|------|------|
| `ps2ics.py` | `.ics` | Apple 日历、Google 日历（Gmail）、Outlook |
| `ps2wakeup.py` | `.csv` | WakeUp 课程表 App |

**无需安装任何第三方库，Python 3.6+ 即可运行。**

---

## 第一步：下载课程表文件

1. 登录 [SIS 系统](https://sis.westlake.edu.cn)
2. 左侧菜单点击 **查看我的课表**
3. 选择顶部 **按课程** 标签页
4. 点击页面**右上角的下载按钮**（见下图红框）

![下载按钮位置](ref.png)

5. 保存文件，默认文件名为 `ps.xls`（实际为 HTML 格式，直接使用即可）

---

## 第二步：生成日历文件

将 `ps.xls` 放到与脚本相同的目录下，打开终端运行：

### Apple 日历 / Google 日历 / Outlook

```bash
python3 ps2ics.py ps.xls
```

生成 `schedule.ics`。

### WakeUp 课程表 App

```bash
python3 ps2wakeup.py ps.xls
```

生成 `wakeup.csv`。

也可以用 `-o` 指定输出路径：

```bash
python3 ps2ics.py ps.xls -o ~/Desktop/westlake.ics
python3 ps2wakeup.py ps.xls -o ~/Desktop/westlake.csv
```

---

## 第三步：导入日历

### Apple 日历

双击 `schedule.ics`，日历会弹出导入确认窗口，点击「导入」即可。

### Google 日历

1. 打开 [Google 日历](https://calendar.google.com) 网页端
2. 右上角齿轮 → **设置**
3. 左侧 **导入和导出** → **导入**
4. 选择 `schedule.ics`，选择导入到的日历，点击「导入」

### Outlook

双击 `schedule.ics`，Outlook 会弹出导入确认窗口，点击「接受」即可。

或手动操作：**文件 → 打开和导出 → 导入/导出 → 导入 iCalendar (.ics)**

### WakeUp 课程表 App

导入前需要先在 WakeUp 中配置西湖大学的节次时间（仅需设置一次）：

**WakeUp → 课表设置 → 节次时间**，共 13 节，按下表填写：

| 节次 | 上课 | 下课 |
|------|------|------|
| 第 1 节 | 08:00 | 08:45 |
| 第 2 节 | 08:50 | 09:35 |
| 第 3 节 | 09:50 | 10:35 |
| 第 4 节 | 10:40 | 11:25 |
| 第 5 节 | 11:30 | 12:15 |
| 第 6 节 | 13:30 | 14:15 |
| 第 7 节 | 14:20 | 15:05 |
| 第 8 节 | 15:10 | 15:55 |
| 第 9 节 | 16:10 | 16:55 |
| 第 10 节 | 17:00 | 17:45 |
| 第 11 节 | 18:30 | 19:15 |
| 第 12 节 | 19:20 | 20:05 |
| 第 13 节 | 20:10 | 20:55 |

节次配置完成后：

1. WakeUp 主界面右上角 → **导入课表**
2. 选择「从 CSV 文件」
3. 选中 `wakeup.csv`，确认导入

---

## 注意事项

- **数据结构与算法设计（EST2005）** 在原始数据中有两条重复的周五记录（教室分别为 H6-303 和 H6-305），导入后可在 WakeUp 或日历中手动删除其中一条
- 时间均以**北京时间（Asia/Shanghai）**写入，跨平台导入不会出现时差问题
- 如有课程导入后时间有误，请检查 SIS 下载的 `ps.xls` 文件是否为当前学期数据
