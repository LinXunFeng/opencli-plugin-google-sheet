# opencli-plugin-google-sheet

基于 `docId` 读取 Google Sheets 的 OpenCLI 插件。

> Language: 中文 | [English](https://github.com/LinXunFeng/opencli-plugin-google-sheet)

## ☕ 请我喝一杯咖啡

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/T6T4JKVRP) [![wechat](https://img.shields.io/static/v1?label=WeChat&message=微信收款码&color=brightgreen&style=for-the-badge&logo=WeChat)](https://cdn.jsdelivr.net/gh/FullStackAction/PicBed@resource20220417121922/image/202303181116760.jpeg)

微信技术交流群请看: [【微信群说明】](https://mp.weixin.qq.com/s/JBbMstn0qW6M71hh-BRKzw)

## 📦 安装

```bash
# 本地路径安装（请替换为你的绝对路径）
opencli plugin install /absolute/path/to/opencli-plugin-google-sheet

# 已安装后更新（本地软链接场景常用）
opencli plugin update google-sheet
```

## 🧩 命令

| 命令 | 说明 |
|---|---|
| `opencli google-sheet sheets --docId <id>` | 列出工作表（`gid`、`title`、`index`） |
| `opencli google-sheet read --docId <id> --sheet <name\|gid>` | 读取指定工作表内容 |

说明：
- `read` 默认输出 `json`（可显式传 `--format json`）。
- `read` 支持 `table/csv/md/plain` 的表格化输出。
- 不传 `--sheet` 时，`read` 会先返回可用工作表列表。

## ❓ 常见问题

### 1) 提示 daemon / extension 未连接

```bash
opencli doctor
opencli daemon stop
```

然后确保 Chrome/Chromium 已打开且 OpenCLI 插件已启用，再重试。

### 2) `sheets` 只有 1 条结果

先验证 `htmlview` 中是否存在多个 gid：

```bash
opencli browser open "https://docs.google.com/spreadsheets/d/<docId>/htmlview?rm=minimal"
opencli browser get html | rg -o 'gid=[0-9]+' | sort -u
```

若能看到多个 gid，插件可直接用 gid 执行读取：

```bash
opencli google-sheet read --docId <docId> --sheet <gid> -f table -v
```

### 3) `AUTH_REQUIRED` / `ACCESS_DENIED`

- 先在同一浏览器里确认该表格可正常打开；
- 保持账号登录状态；
- 再执行命令重试。

## 🛠️ 开发

```bash
# 进入插件目录
cd /absolute/path/to/google-sheet

# 本地安装（软链接）
opencli plugin install "$PWD"

# 查看命令是否注册成功
opencli list | grep "google-sheet"

# 运行测试
npm test
```

## 🖨 关于我

- GitHub: [https://github.com/LinXunFeng](https://github.com/LinXunFeng)
- Email: [linxunfeng@yeah.net](mailto:linxunfeng@yeah.net)
- Blogs: 
  - 全栈行动: [https://fullstackaction.com](https://fullstackaction.com)
  - 掘金: [https://juejin.cn/user/1820446984512392](https://juejin.cn/user/1820446984512392) 

<img height="267.5" width="481.5" src="https://github.com/LinXunFeng/LinXunFeng/raw/master/static/img/FSAQR.png"/>
