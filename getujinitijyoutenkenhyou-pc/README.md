# 月次日常点検 Webアプリ（MVP）

## 起動方法
1. Node.js 18 以上を用意
2. プロジェクト直下で実行

```bash
npm start
```

3. ブラウザで `http://localhost:3000` を開く

## GitHub Pagesで開く
- ルートの `index.html` から `src/index.html` へ遷移します。
- Pagesでは `server/server.js` が動かないため、保存先はブラウザの `localStorage` になります。
- 同じブラウザ・同じ端末でのみ保存データを再読込できます。

## 実装済み
- 点検セルクリックで `レ -> ☓ -> ▲ -> 空欄`
- 運行管理者印: 岸田
- 整備管理者印: 若本
- 月・車番・運転者キーで保存/読込

## ファイル
- `server/server.js`: API + 静的配信
- `src/index.html`: 画面
- `src/main.js`: 画面ロジック
- `src/styles.css`: スタイル
- `docs/roadmap-1month.md`: 1か月計画
