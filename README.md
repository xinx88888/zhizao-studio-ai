# 智造工作室 AI 工作台

智造工作室的 AI 协作平台，集成任务管理、财务追踪、内容脚本生成、联网调研、编程助手等功能。

## 功能特性

- AI 智能顾问（Ollama 本地模型）
- 任务计划与追踪
- 财务管理（Excel 集成）
- 内容脚本库
- 联网市场调研
- 团队管理（三级权限：管理员/编辑/只读）
- 系统设置与主题切换

## 技术栈

- 后端：Flask + SQLite/PostgreSQL
- 前端：原生 HTML/JS + CSS
- AI：Ollama (本地) / OpenAI API (云端)
- 部署：Render / Gunicorn

## 本地开发

```bash
pip install -r requirements.txt
python studio_app.py
```

访问 http://localhost:18080

默认账号：
- 管理员：熊科瑞 / 智造2026
- 编辑：韦硕 / 智造2026

## 环境变量

| 变量 | 说明 | 默认值 |
|------|------|--------|
| DATABASE_URL | PostgreSQL 连接串 | 空（使用 SQLite）|
| SECRET_KEY | Flask 密钥 | studio_v4_secret_2024 |
| FINANCE_EXCEL | 财务 Excel 路径 | E:\Desktop\财务系统-收入支出管理1.xlsx |
| PORT | HTTP 端口 | 18080 |

## 部署

### Render

1. 访问 https://dashboard.render.com
2. 点击 New -> PostgreSQL -> 选择 Free 方案
3. 创建完成后，复制 Internal Database URL
4. 点击 New -> Web Service
5. 选择 GitHub 仓库: xinx88888/zhizao-studio-ai
6. 配置:
   - Build Command: pip install -r requirements.txt
   - Start Command: gunicorn -w 4 -b 0.0.0.0:$PORT studio_app:app
7. 添加环境变量:
   - DATABASE_URL=<你的 PostgreSQL URL>
   - SECRET_KEY=<随机字符串，如: zhizao2024secret>
   - FINANCE_EXCEL=/tmp/finance.xlsx
8. 点击 Deploy

## 许可证

MIT
