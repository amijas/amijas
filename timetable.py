from datetime import date, timedelta, datetime
import jpholiday
from ortools.sat.python import cp_model
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from icalendar import Calendar, Event
import pytz
import csv  # ファイルの先頭に追加

# 1. メンバーリスト
members = [
    "Abe",
    "Tanifuji",
    "Tanaka",
    "Hishikawa",
    "Tokimura",
    "Sajima",
    "Komai",
    "Aqsa",
    "Tsujimoto",
    "Kato",
    "Huang",
]

# 2. 特定の制約
member_constraints = {
    "Abe": ["火曜日"],
    "Tanifuji": ["木曜日", "金曜日"],
    "Hishikawa": ["金曜日"],
    "Tokimura": ["水曜日", "木曜日"],
    "Sajima": ["水曜日"],
    "Tsujimoto": ["木曜日"],
    "Huang": ["木曜日", "金曜日"],
}


# 3. 特定の月の平日を取得
def get_weekdays_of_specific_month(year, month):
    first_day_of_month = date(year, month, 1)
    current_day = first_day_of_month
    weekdays = []

    # 日本語の曜日名に変換するための辞書
    weekday_names = {
        0: "月曜日",
        1: "火曜日",
        2: "水曜日",
        3: "木曜日",
        4: "金曜日",
        5: "土曜日",
        6: "日曜日",
    }

    while current_day.month == month:
        # 月曜日(0)、土曜日(5)、日曜日(6)、祝日を除外
        if (
            current_day.weekday() > 0  # 月曜日を除外
            and current_day.weekday() < 5  # 土日を除外
            and not jpholiday.is_holiday(current_day)  # 祝日を除外
        ):
            # 日本語の曜日名を使用
            weekdays.append((current_day, weekday_names[current_day.weekday()]))
        current_day += timedelta(days=1)
    return weekdays


# 4. スケジューリングの最適化
def create_schedule(members, weekdays, member_constraints):
    # 月曜日を除外したweekdaysを作成
    weekdays = [(day, name) for day, name in weekdays if day.weekday() != 0]

    num_days = len(weekdays)
    num_members = len(members)

    # モデルの作成
    model = cp_model.CpModel()

    # 変数の定義
    schedule = {}
    for d in range(num_days):
        for m in range(num_members):
            schedule[(d, m)] = model.NewBoolVar(f"day_{d}_member_{m}")

    # 制約1: 各日には1人のメンバーのみ割り当て
    for d in range(num_days):
        model.Add(sum(schedule[(d, m)] for m in range(num_members)) == 1)

    # 制約2: 各メンバーの担当日数を均等に（重み付けあり）
    weighted_days = 0
    for _, (_, day_name) in enumerate(weekdays):
        # 水曜日は2日分としてカウント
        weighted_days += 2 if day_name == "水曜日" else 1

    min_weighted_days_per_member = weighted_days // num_members
    max_weighted_days_per_member = min_weighted_days_per_member + 1

    for m, member in enumerate(members):
        weighted_sum = sum(
            2 * schedule[(d, m)] if weekdays[d][1] == "水曜日" else schedule[(d, m)]
            for d in range(num_days)
        )

        if member == "Abe":
            # Abeは1回のみ（水曜日の場合は2としてカウント）
            model.Add(weighted_sum >= 1)
            model.Add(weighted_sum <= 2)  # 水曜日の場合は2までOK
        else:
            # 他のメンバーは重み付けした日数で均等化
            model.Add(weighted_sum >= min_weighted_days_per_member)
            model.Add(weighted_sum <= max_weighted_days_per_member)

    # 制約3: 同じメンバーが連続で割り当てれない
    for d in range(1, num_days):
        for m in range(num_members):
            model.Add(schedule[(d, m)] + schedule[(d - 1, m)] <= 1)

    # 制約4: 特定の曜日の制約を適用（修正）
    for d, (_, day_name) in enumerate(weekdays):
        for m, member in enumerate(members):
            # メンバーが特定の曜日に入れない場合は0を設定
            if member in member_constraints and day_name in member_constraints[member]:
                model.Add(schedule[(d, m)] == 0)  # その曜日には入れない

    # 制約5: 待機の間隔は5日以上空ける
    for d in range(num_days):
        for m in range(num_members):
            for offset in range(1, 5):
                if d + offset < num_days:
                    model.Add(schedule[(d, m)] + schedule[(d + offset, m)] <= 1)

    # ソルバーの設定
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # スケジュール結果の取得
    if status == cp_model.OPTIMAL:
        result = []
        for d, (day, day_name) in enumerate(weekdays):
            for m in range(num_members):
                if solver.Value(schedule[(d, m)]) == 1:
                    result.append((day, day_name, members[m]))
        return result
    else:
        return None


class TimetableApp:
    def __init__(self, root):
        self.root = root
        self.root.title("当番表作成アプリ")

        # ウィンドウの最小サイズを設定
        self.root.minsize(400, 600)  # 幅400px、高さ600px

        # メインフレームの作成（スクロール可能にする）
        main_frame = ttk.Frame(root)
        main_frame.grid(row=0, column=0, sticky="nsew")

        # 年の入力
        ttk.Label(main_frame, text="年:").grid(row=0, column=0, padx=5, pady=5)
        self.year_var = tk.StringVar(value="2025")
        self.year_entry = ttk.Entry(main_frame, textvariable=self.year_var)
        self.year_entry.grid(row=0, column=1, padx=5, pady=5)

        # 月の入力
        ttk.Label(main_frame, text="月:").grid(row=1, column=0, padx=5, pady=5)
        self.month_var = tk.StringVar(value="2")
        self.month_entry = ttk.Entry(main_frame, textvariable=self.month_var)
        self.month_entry.grid(row=1, column=1, padx=5, pady=5)

        # 実行ボタン
        ttk.Button(main_frame, text="当番表を作成", command=self.create_timetable).grid(
            row=2, column=0, columnspan=2, pady=20
        )

        # 結果表示エリア
        self.result_text = tk.Text(main_frame, height=20, width=50)
        self.result_text.grid(row=3, column=0, columnspan=2, padx=5, pady=5)

        # ボタンを横に並べるためのフレーム
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=10, padx=5)

        # iCalendarエクスポートボタン
        ttk.Button(
            button_frame,
            text="iCalendarファイルとして保存",
            command=self.export_to_ical,
        ).pack(side=tk.LEFT, padx=5)

        # CSVエクスポートボタン
        ttk.Button(
            button_frame, text="CSVファイルとして保存", command=self.export_to_csv
        ).pack(side=tk.LEFT, padx=5)

        # グリッドの設定
        root.grid_rowconfigure(0, weight=1)
        root.grid_columnconfigure(0, weight=1)

    def create_timetable(self):
        try:
            year = int(self.year_var.get())
            month = int(self.month_var.get())

            weekdays = get_weekdays_of_specific_month(year, month)
            self.schedule = create_schedule(members, weekdays, member_constraints)

            # 結果の表示
            self.result_text.delete(1.0, tk.END)
            if self.schedule:
                self.result_text.insert(tk.END, "スケジュール:\n")
                for day, day_name, member in self.schedule:
                    if day.weekday() != 0:  # 0は月曜日
                        self.result_text.insert(
                            tk.END, f"{day} ({day_name}): {member}\n"
                        )

                # 当直回数の集計と表示
                duty_counts = {}
                for _, _, member in self.schedule:
                    duty_counts[member] = duty_counts.get(member, 0) + 1

                self.result_text.insert(tk.END, "\n当直回数:\n")
                for member in sorted(members):  # メンバーリストの順番で表示
                    count = duty_counts.get(member, 0)
                    self.result_text.insert(tk.END, f"{member}: {count}回\n")
            else:
                self.result_text.insert(tk.END, "スケジュールを作成できませんでした。")

        except ValueError:
            messagebox.showerror("エラー", "正しい年月を入力してください")
        except Exception as e:
            messagebox.showerror("エラー", f"エラーが発生しました: {str(e)}")

    def export_to_ical(self):
        if not hasattr(self, "schedule") or not self.schedule:
            messagebox.showerror("エラー", "先にスケジュールを作成してください")
            return

        try:
            # カレンダーオブジェクトの作成
            cal = Calendar()
            cal.add("prodid", "-//当番表作成アプリ//example.com//")
            cal.add("version", "2.0")

            # タイムゾーンの設定
            japan_tz = pytz.timezone("Asia/Tokyo")

            # 各当番をイベントとして追加
            for day, day_name, member in self.schedule:
                if day.weekday() != 0:  # 月曜日を除外
                    event = Event()
                    event.add("summary", f"当番: {member}")

                    # 終日イベントとして設定
                    event.add("dtstart", day, parameters={"VALUE": "DATE"})
                    # 終了日は開始日の翌日
                    event.add(
                        "dtend", day + timedelta(days=1), parameters={"VALUE": "DATE"}
                    )

                    event.add("description", f"{day_name}の当番担当")
                    # 終日イベントであることを示すプロパティ
                    event["X-MICROSOFT-CDO-ALLDAYEVENT"] = "TRUE"
                    event["TRANSP"] = "TRANSPARENT"  # 予定あり/なしの表示用
                    cal.add_component(event)

            # 保存先の選択
            file_path = filedialog.asksaveasfilename(
                defaultextension=".ics",
                filetypes=[("iCalendarファイル", "*.ics")],
                initialfile=f"当番表_{self.year_var.get()}_{self.month_var.get()}.ics",
            )

            if file_path:
                with open(file_path, "wb") as f:
                    f.write(cal.to_ical())
                messagebox.showinfo("成功", "iCalendarファイルを保存しました")

        except Exception as e:
            messagebox.showerror(
                "エラー", f"ファイルの保存中にエラーが発生しました: {str(e)}"
            )

    def export_to_csv(self):
        if not hasattr(self, "schedule") or not self.schedule:
            messagebox.showerror("エラー", "先にスケジュールを作成してください")
            return

        try:
            # 保存先の選択
            file_path = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSVファイル", "*.csv")],
                initialfile=f"当番表_{self.year_var.get()}_{self.month_var.get()}.csv",
            )

            if file_path:
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    # ヘッダーを書き込む
                    writer.writerow(
                        [
                            "Subject",
                            "Start Date",
                            "Start Time",
                            "End Date",
                            "End Time",
                            "All Day Event",
                            "Description",
                        ]
                    )

                    # データを書き込む
                    for day, day_name, member in self.schedule:
                        if day.weekday() != 0:  # 月曜日を除外
                            writer.writerow(
                                [
                                    f"当番:{member}",  # Subject
                                    day.strftime("%Y/%m/%d"),  # Start Date
                                    "0:00",  # Start Time
                                    day.strftime("%Y/%m/%d"),  # End Date
                                    "23:59",  # End Time
                                    "TRUE",  # All Day Event
                                    f"{day_name}の当番担当",  # Description
                                ]
                            )

                messagebox.showinfo("成功", "CSVファイルを保存しました")

        except Exception as e:
            messagebox.showerror(
                "エラー", f"ファイルの保存中にエラーが発生しました: {str(e)}"
            )


# GUIの起動
if __name__ == "__main__":
    root = tk.Tk()
    app = TimetableApp(root)
    root.mainloop()
