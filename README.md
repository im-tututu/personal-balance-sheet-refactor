# GAS Local Workflow

当前项目绑定的 Apps Script 本地开发，推荐用下面这套节奏。

## 1. 安装依赖

本仓库没有运行时依赖，只需要本机有：

- Node.js
- `clasp`

```bash
npm install
```

如果本机还没有 `clasp`：

```bash
npm install -g @google/clasp
```

## 2. 切换脚本环境

生产脚本配置已经放在：

- `.clasp.prod.json`

开发脚本请自己复制一份：

```bash
cp .clasp.dev.json.example .clasp.dev.json
```

把 `.clasp.dev.json` 里的 `scriptId` 改成你的开发脚本 ID。

切环境：

```bash
npm run gas:env:prod
npm run gas:env:dev
```

## 3. 日常开发

建议开两个终端。

终端 1：自动推送

```bash
npm run gas:watch
```

终端 2：实时日志

```bash
npm run gas:logs
```

这样保存本地文件后会自动 push，不用每次手工执行 `clasp push`。

如果提示 `Invalid scriptId in .clasp.json` 或 `Request contains an invalid argument.`，
通常是因为当前 `.clasp.json` 还是占位符，先执行：

```bash
npm run gas:env:prod
```

或者先准备好 `.clasp.dev.json`，再执行：

```bash
npm run gas:env:dev
```

## 4. 本地先跑基础测试

语法检查：

```bash
npm run test:syntax
```

纯计算 smoke test：

```bash
npm run test:calc
```

全部一起跑：

```bash
npm test
```

## 5. 常用命令

```bash
npm run gas:push
npm run gas:open
npm run gas:status
```

## 6. 推荐测试方式

最省时间的方式不是直接在生产表上测，而是：

1. 复制一份开发版 Spreadsheet
2. 绑定一个开发版 Apps Script
3. 本地切到 `dev`
4. 开 `gas:watch` + `gas:logs`
5. 在开发表里执行测试函数

这样保留真实表格环境，同时避免污染生产数据。
