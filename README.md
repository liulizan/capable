# Capable（可培）— AI 能力转化引擎（离线版）

把优质内容（文章、笔记、教程）转化为可练习、可落地的个人能力。基于 TextRank + TF-IDF 本地智能分析引擎，无需任何外部 API。

## 项目结构

```
capable/
├── package.json          # 项目配置与依赖
├── server.js             # 后端（Express + 本地分析引擎）
├── .env                  # 环境变量（可选，存放 PORT 等配置）
├── public/
│   ├── index.html        # 前端单页
│   ├── script.js         # 前端交互逻辑
│   └── style.css         # 样式
└── README.md
```

## 快速开始

### 1. 安装依赖

```bash
cd capable
npm install
```

### 2. 配置（可选）

在项目根目录创建 `.env` 文件：

```
PORT=3000
```

不创建 `.env` 也可直接使用，默认监听 3000 端口。

### 3. 启动服务

```bash
npm start
```

或使用开发模式（自动重载）：

```bash
npm run dev
```

### 4. 打开浏览器

访问 `http://localhost:3000`

## 使用流程

1. （可选）选择你的身份/角色和学习目标
2. 粘贴要转化的文章或笔记内容（支持 .docx / .xlsx / .md 文件上传和链接抓取）
3. 选择输出模式（总结重点 / 提取信息 / 批判分析 / 生成练习卡片）
4. 点击 **"开始能力转化"** → 输出能力卡片 + 行动脚本 + 自测题 + 角色扮演
5. 在角色扮演区域回复 → AI 给出点评 + 反馈

## 技术栈

| 层 | 技术 |
|---|---|
| 后端 | Node.js + Express |
| 前端 | 原生 HTML/CSS/JS（无框架） |
| 分析 | TextRank + TF-IDF 本地智能引擎 |
| 渲染 | marked.js（CDN） |

## API 接口

### POST /api/transform

转化内容，支持 4 种模式（可多选）：总结重点、提取信息、批判分析、生成练习卡片。

请求体：
```json
{
  "content": "要转化的文章内容（必填）",
  "userProfile": "身份/目标（可选）",
  "modes": ["summarize", "capability"]
}
```

### POST /api/roleplay

角色扮演点评 + 收集反馈。

请求体：
```json
{
  "conversationHistory": [...],
  "userReply": "用户在角色扮演中的回复",
  "sessionId": "会话ID"
}
```

### POST /api/parse-file

上传 .docx / .xlsx / .md / .txt 文件，返回解析后的纯文本。

### POST /api/fetch-url

抓取网页链接内容，返回提取的纯文本。

## 安全说明

- 所有分析在本地完成，不上传任何数据到外部服务器
- 会话数据仅存储在服务端内存，1 小时后自动过期
- 支持文件上传和链接抓取，无需任何第三方 API Key
