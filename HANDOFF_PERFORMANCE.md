# 高速化作業 引継ぎメモ

- 更新日: 2026-08-04
- 本番URL: https://aoyama-kokyaku.streamlit.app/
- GitHub: JIRO-AOYAMA/aoyama-kokyaku
- 本番基準コミット: `4301466016a8c0483800fb64ebdc5d675373d931`
- 作業ブランチ: `perf/phase1-measurement-20260804`
- 作業フォルダ: `C:\Dev\aoyama-kokyaku\github-main`

## 現在の状態

- GitHubの最新mainから専用ブランチを作成済み。
- 本番の`main`と本番Streamlitは未変更。
- `app.py`の変更はまだない。
- OneDrive内の旧`.py`は使用しない。

## 次に行うこと

1. 本番機能を変えない速度計測をテストブランチへ追加する。
2. Dropbox、Excel解析、Supabase、主要画面の処理時間を確認する。
3. 既存機能を確認してから、最も効果の大きい改善を1件だけ実施する。

## 安全ルール

- 一度に1種類だけ変更する。
- 未確認の変更をmainへ入れない。
- Excel・Supabaseの保存処理は最初の高速化では変更しない。
- 秘密鍵やパスワードをこのファイルへ記載しない。
- 各作業後に、このメモとGitコミットを更新する。

