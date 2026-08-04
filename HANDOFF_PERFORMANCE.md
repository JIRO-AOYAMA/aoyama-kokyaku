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
- テストブランチの`app.py`へ、`perf=1`の時だけ表示される速度診断を追加済み。
- 通常のURLでは診断処理は無効で、Excel・Supabaseの保存処理は未変更。
- テスト環境の起動時に発生した配車モジュールの再読込エラーに対し、`dispatch_filters.py`からモジュール間の再importをなくし、同じ正規化処理をファイル内へ配置。
- 新規環境のStreamlit 1.60.0でCookie準備待ちの白画面になったため、テストブランチだけ直前版`1.59.2`へ固定して比較中。
- The isolated app `aoyama-kokyaku-cszjcu4w4rmykjojby9wxm.streamlit.app` stayed blank even after `app.py` was replaced with a minimal three-widget page and the app was rebooted. This rules out the customer app code and points to a broken Streamlit app instance/deployment.
- The full diagnostic `app.py` was restored exactly from commit `cd0150f` in restore commit `43c56c8`.

## Private performance app (2026-08-04)

- New test app: `https://aoyama-kokyaku-qwlh5ekeys6bgnhjovzawx.streamlit.app/`
- Repository/branch/file: `JIRO-AOYAMA/aoyama-kokyaku` / `perf/phase1-measurement-20260804` / `app.py`
- Python: `3.12`
- The user turned off `Make this app public` before the test-only login gate was deployed.
- Commit `db0de3e` adds a test-only login gate. It is enabled only when all three conditions are true:
  1. The request host exactly matches the new test app host.
  2. The URL contains `perf=1`.
  3. The test app has the top-level Streamlit Secret `PERFORMANCE_TEST_MODE = true`.
- Synthetic test claims are not written to the Supabase Microsoft login history.
- Production `main` and the production Streamlit app were not changed.

### Immediate next step

1. Add `PERFORMANCE_TEST_MODE = true` at the very top of the new private test app's Secrets, before the first TOML section header such as `[auth]`.
2. Save the Secrets.
3. Open `https://aoyama-kokyaku-qwlh5ekeys6bgnhjovzawx.streamlit.app/?perf=1`.
4. Confirm the home screen and the private performance-test warning appear without Microsoft login.
5. Measure the baseline before making any performance optimization.
- OneDrive内の旧`.py`は使用しない。

## 次に行うこと

1. Do not reuse the blank Streamlit app instance. Create a fresh test app from `perf/phase1-measurement-20260804`, then verify that the Microsoft login screen appears before copying production Secrets.
2. URL末尾に`&perf=1`を付け、Dropbox、Excel解析、Supabase、主要画面の処理時間を確認する。
3. 既存機能を確認してから、最も効果の大きい改善を1件だけ実施する。

## 安全ルール

- 一度に1種類だけ変更する。
- 未確認の変更をmainへ入れない。
- Excel・Supabaseの保存処理は最初の高速化では変更しない。
- 秘密鍵やパスワードをこのファイルへ記載しない。
- 各作業後に、このメモとGitコミットを更新する。
