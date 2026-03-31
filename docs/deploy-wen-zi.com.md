# `wen-zi.com` Docker / GitHub Actions 部署手顺

このプロジェクトは Slack `Socket Mode` の常駐 bot なので、公開 HTTP エンドポイントは不要です。
本番では Docker コンテナ 1 個だけを常駐させ、GitHub Actions からソースを転送し、サーバー側で Docker build して自動デプロイします。

## 現在のサーバー方針

- `8088` の `vitality-server-backend` はそのまま残す
- 旧 `expenses-agent-api` と `expenses-agent-slack-bot` は新デプロイ時に停止・削除する
- 新 bot は `/opt/slack-excel-bot` 配下で Docker Compose 管理する
- 80 番ポートは新 bot では使わない

## サーバー構成

デプロイ先ディレクトリ:

```text
/opt/slack-excel-bot
  docker-compose.prod.yml
  deploy.sh
  .env
  .release.json
  .data/
```

Compose は 1 サービスだけです。

```yaml
services:
  slack-excel-bot:
    image: ${IMAGE_URI:?IMAGE_URI is required}
    container_name: slack-excel-bot
    env_file:
      - .env
    volumes:
      - ${HOST_DATA_DIR:-./.data}:/app/.data
    restart: unless-stopped
```

## GitHub Actions

### CI

`.github/workflows/ci.yml`

- `push`
- `pull_request`

実行内容:

1. Python 3.11 をセットアップ
2. `pip install -e ".[dev]"`
3. `pytest -q`
4. `docker build`

### CD

`.github/workflows/cd.yml`

トリガー:

- `v*` 形式の tag push

前提条件:

- tag のコミットが `origin/main` から到達可能であること

実行内容:

1. AWS 認証
2. ECR ログイン
3. tag 対応のソースを `release.tar.gz` として作成
4. SSH で `wen-zi.com` に接続
5. `/opt/slack-excel-bot/.env` を更新
6. サーバーでソースを展開し `docker build -t slack-excel-bot:<tag>` を実行
7. `/opt/slack-excel-bot/deploy.sh` を実行
8. 新コンテナ起動後に旧 `expenses-agent-*` コンテナを停止・削除

## GitHub Secrets

GitHub リポジトリ側で以下を設定します。

- `DEPLOY_HOST`
- `DEPLOY_USER`
- `DEPLOY_SSH_KEY`
- `SERVER_ENV_FILE`

推奨値:

- `DEPLOY_HOST=wen-zi.com`
- `DEPLOY_USER=ec2-user`

`DEPLOY_SSH_KEY` には `/Users/srwenzi/Downloads/JapanServerKey.pem` の中身をそのまま登録します。

`SERVER_ENV_FILE` には本番用 `.env` の全文を登録します。最低限の例:

```dotenv
SLACK_BOT_TOKEN=xoxb-...
SLACK_APP_TOKEN=xapp-...
OPENAI_API_KEY=sk-...
EXPENSES_EKISPERT_API_TOKEN=...
OPENAI_MODEL=gpt-5.4
STORAGE_DIR=/app/.data
DEFAULT_EMPLOYEE_NAME=山田太郎
DEFAULT_EMPLOYEE_ID=0001
DEFAULT_DEPARTMENT=開発本部
DEFAULT_DEPARTMENT_CODE=50
DEFAULT_WORK_GRADE=1
DEFAULT_CLOCK_IN=09:00
DEFAULT_CLOCK_OUT=18:00
MAX_CONCURRENT_REQUESTS=50
```

## デプロイフロー

初回:

1. GitHub に新規リポジトリを作成する
2. このローカルリポジトリへ `origin` を追加する
3. `main` を push する
4. GitHub の default branch を `main` に設定する
5. Secrets を登録する
通常リリース:

```bash
git checkout main
git pull
git tag v0.1.0
git push origin main --tags
```

これで GitHub Actions が自動的に:

- tag のソースを `wen-zi.com` に転送
- サーバーで `slack-excel-bot:v0.1.0` を build
- `wen-zi.com` 上の `slack-excel-bot` を更新
- 旧 `expenses-agent-api` / `expenses-agent-slack-bot` を停止

## サーバー側デプロイスクリプトの責務

`deploy/deploy.sh` は冪等に動くようにしてあります。

- `/opt/slack-excel-bot` と `.data` を作成
- サーバーで build 済みのローカル image を使って `docker compose up -d --remove-orphans`
- `docker compose up -d --remove-orphans`
- 旧コンテナ停止と削除
- `.release.json` 更新

## デプロイ後チェック

- `docker ps` に `slack-excel-bot` が出る
- `docker ps` に `expenses-agent-api` と `expenses-agent-slack-bot` が出ない
- `docker ps` に `vitality-server-backend` は残る
- Slack DM に bot が応答する
- Excel が upload される
- 交通費機能で Ekispert が使える
