# 精算くん

这是一个名为 `精算くん` 的 Slack chatbot 项目，用来在 1 对 1 会话里读取用户输入的文字和图片，交给 OpenAI LLM 判断意图并调用 Excel 生成工具，最后把生成好的文件上传回 Slack 会话。

当前项目只保留 `Socket Mode`，不再依赖 HTTP Events 回调。

当前内置 3 个 Excel 工具：

- `考勤表`
- `交通费精算表`
- `个人报销计算表`

项目只复用了旧项目中与 Excel 模板填充直接相关的部分：模板文件、模板映射思路和 writer 行为。Slack 流程、OpenAI 调用和整体架构都是在当前目录重新实现的。

## 能力范围

- 监听 Slack 私聊 `message.im` 事件
- 读取用户消息文字
- 读取用户上传的图片，并把图片内容传给 OpenAI 模型
- 让模型通过 tool calling 选择 3 个 Excel 工具之一
- 工具把 JSON 参数转换成 Excel 文件
- 生成完成后把文件上传回原对话

## 运行方式

1. 安装依赖

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[dev]"


cd /Users/srwenzi/workspace/ldstAgent
source .venv/bin/activate
python -m slack_excel_bot.app

```

2. 配置环境变量

```bash
cp .env.example .env
```

至少需要配置：

- `SLACK_BOT_TOKEN`
- `SLACK_APP_TOKEN`
- `OPENAI_API_KEY` 或 `EXPENSES_LLM_API_KEY`

如果你的模板要求员工编号、姓名、部门这些固定字段，建议同时配置：

- `DEFAULT_EMPLOYEE_NAME`
- `DEFAULT_EMPLOYEE_ID`
- `DEFAULT_DEPARTMENT`
- `DEFAULT_DEPARTMENT_CODE`

这样用户只说“帮我做一个三月全勤的表”时，机器人也能直接补齐模板里必须的头部信息。

3. 启动服务

```bash
python -m slack_excel_bot.app
```

## Slack App 配置建议

- 开启 `Socket Mode`
- App-level token 需要 `connections:write`
- 订阅事件：`message.im`
- Bot Token Scopes 至少包含：
  - `chat:write`
  - `files:write`
  - `im:history`
  - `im:read`
  - `files:read`

## Excel 文件输出逻辑

当前不是让你自己再额外 `serve` 文件。

实际流程是：

1. 后端先在本地 `STORAGE_DIR/drafts/` 生成 `.xlsx`
2. 然后直接调用 Slack `files_upload_v2`
3. 文件会被上传进当前用户私聊
4. 再发送一条文字回复

所以现状是“本地生成，直接上传到 Slack”，不是“后端生成一个下载链接让 Slack 去拉”。

## OpenAI 设计说明

本项目使用 OpenAI 的 `Responses API` 做多模态输入和 tool calling：

- 文本通过 `input_text` 发送
- 图片通过 `input_image` 发送
- Excel 生成器作为函数工具暴露给模型

因此像“帮我根据这张交通记录截图做交通费精算表”这种请求，模型可以直接看到截图并决定调用 `交通费精算表` 工具。

## 目录结构

```text
src/slack_excel_bot/
  app.py
  config.py
  slack_bot.py
  socket_mode.py
  openai_agent.py
  excel_tools.py
  excel_writer.py
  template_loader.py
  template_schema.py
  tool_schemas.py
  templates/
    files/
    mappings/
    registry.json
tests/
```

## Docker 运行

本项目生产环境使用单容器 Docker 运行，不暴露公网端口。

构建镜像：

```bash
docker build -t slack-excel-bot:local .
```

本地用 `.env` 启动：

```bash
docker run --rm \
  --env-file .env \
  -v "$PWD/.data:/app/.data" \
  slack-excel-bot:local
```

## CI / CD

仓库包含两条 GitHub Actions：

- `CI`：每次 `push` 和 `pull_request` 都会运行 `pytest -q` 并校验 Docker build。
- `CD`：只有给 `main` 上的提交打 `v*` tag 并推送时才会发布。

CD 发布流程：

1. 校验 tag 指向的提交可从 `origin/main` 到达
2. 打包当前 tag 对应源码并上传到 `wen-zi.com`
3. 在服务器本地执行 `docker build`
4. 更新 `/opt/slack-excel-bot/.env`
5. 运行部署脚本重建 `slack-excel-bot` 容器
6. 停掉旧项目 `expenses-agent-api` 和 `expenses-agent-slack-bot`

GitHub Secrets 约定：

- `DEPLOY_HOST`
- `DEPLOY_USER`
- `DEPLOY_SSH_KEY`
- `SERVER_ENV_FILE`

`SERVER_ENV_FILE` 至少应包含：

```dotenv
SLACK_BOT_TOKEN=xoxb-...
SLACK_APP_TOKEN=xapp-...
OPENAI_API_KEY=sk-...
EXPENSES_EKISPERT_API_TOKEN=...
OPENAI_MODEL=gpt-5.4
STORAGE_DIR=/app/.data
MAX_CONCURRENT_REQUESTS=50
```

## 部署文档

`wen-zi.com` 的 Docker + GitHub Actions 部署步骤见：

- `docs/deploy-wen-zi.com.md`
