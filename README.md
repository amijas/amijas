# 当番表作成アプリ

## 概要
当番表を自動で作成するPythonアプリケーションです。Google OR-Toolsを使用して最適な当番スケジュールを生成し、iCalendarまたはCSV形式でエクスポートできます。

## 特徴
- 直感的なGUIインターフェース
- 月曜日・土日・祝日を自動的に除外
- メンバーごとの制約設定（特定の曜日に入れない等）
- 水曜日の負荷を考慮（2日分としてカウント）
- 特定メンバー（Abe）の当直回数を1回以下に制限
- 連続当直の防止（5日以上の間隔を確保）
- 当直回数の均等化
 
